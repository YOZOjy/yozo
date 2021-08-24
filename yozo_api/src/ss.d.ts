declare namespace excel {
    export interface AboveAverage {
        setAboveBelow(s: int): void
        getAboveBelow(): int
    }
    export interface AboveAverage {
        setAboveBelow(s: int): void
        getAboveBelow(): int
    }
    export interface AddIn {
        getName(): string
        getPath(): string
        getFullName(): string
        setInstalled(install: boolean): void
        isInstalled(): boolean
        getProgID(): string
        getCLSID(): string
    }
    export interface AddIn {
        getName(): string
        getPath(): string
        getFullName(): string
        setInstalled(install: boolean): void
        isInstalled(): boolean
        getProgID(): string
        getCLSID(): string
    }
    export interface AddIns {
        add(fileName: string, copyFile: boolean): word.AddIn
        item(index: any): word.AddIn
        getCount(): int
    }
    export interface AddIns {
        add(fileName: string, copyFile: boolean): word.AddIn
        item(index: any): word.AddIn
        getCount(): int
    }
    export interface Adjustments {
        item(index: int): float
        getCount(): int
    }
    export interface Adjustments {
        item(index: int): float
        getCount(): int
    }
    export interface AllowEditRange {
        delete(): void
        setTitle(title: string): void
        getTitle(): string
        getRange(): excel.Range
        unprotect(password: string): void
        setRange(range: excel.Range): void
        getUsers(): excel.UserAccessList
        changePassword(password: string): void
    }
    export interface AllowEditRange {
        delete(): void
        setTitle(title: string): void
        getTitle(): string
        getRange(): excel.Range
        unprotect(password: string): void
        setRange(range: excel.Range): void
        getUsers(): excel.UserAccessList
        changePassword(password: string): void
    }
    export interface AllowEditRanges {
        add(title: string, range: excel.Range, password: string): excel.AllowEditRange
        item(obj: any): excel.AllowEditRange
        getCount(): int
    }
    export interface AllowEditRanges {
        add(title: string, range: excel.Range, password: string): excel.AllowEditRange
        item(obj: any): excel.AllowEditRange
        getCount(): int
    }
    export interface Application {
        getName(): string
        getValue(): string
        close(saveChanges: int): void
        getPath(): string
        getPathSeparator(): string
        closeAll(): int
        help(helpFile: any, helpContextID: any): void
        union(arg1: excel.Range, arg2: excel.Range, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): excel.Range
        setCursor(cursor: int): void
        getCursor(): int
        getBuild(): int
        getWidth(): double
        getLeft(): double
        setLeft(left: double): void
        setTop(top: double): void
        getTop(): double
        setWidth(width: double): void
        DDEPoke(Channel: number, Item: string, Data: string): void
        getMainControl(): any
        getCommandBars(): office.CommandBars
        setActivePrinter(activePrinter: string): void
        getAutomationSecurity(): int
        setAutomationSecurity(AutomationSecurity: int): void
        isArbitraryXMLSupportAvailable(): boolean
        getActivePrinter(): string
        getActiveWindow(): excel.Window
        getWindows(): excel.Windows
        getAddIns(): excel.AddIns
        getAnswerWizard(): office.AnswerWizard
        getAssistant(): office.Assistant
        getAutoCorrect(): excel.AutoCorrect
        setCaption(caption: string): void
        getCaption(): string
        getCOMAddIns(): office.COMAddIns
        getOptions(): office.AbstractOptions
        getOptions(): excel.Options
        getDialogs(): excel.Dialogs
        setDisplayAlerts(displayAlerts: boolean): void
        getDefaultSaveFormat(): int
        setDefaultSaveFormat(defaultSaveFormat: int): void
        isDisplayRecentFiles(): boolean
        setDisplayRecentFiles(displayRecentFiles: boolean): void
        isDisplayScrollBars(): boolean
        setDisplayScrollBars(displayScrollBars: boolean): void
        isDisplayStatusBar(): boolean
        setDisplayStatusBar(displayStatusBar: boolean): void
        getEnableCancelKey(): int
        setEnableCancelKey(enableCancelKey: int): void
        getFeatureInstall(): int
        setFeatureInstall(featureInstall: int): void
        getFileConverters(index1: any, index2: any): any
        getFileDialog(fileDialogType: int): office.FileDialog
        getFileSearch(): office.FileSearch
        getHeight(): double
        setHeight(height: double): void
        getInternational(index: any): any
        isMouseAvailable(): boolean
        getRecentFiles(): excel.RecentFiles
        isScreenUpdating(): boolean
        isMathCoprocessorAvailable(): boolean
        setScreenUpdating(screenUpdating: boolean): void
        isShowStartupDialog(): boolean
        setShowStartupDialog(showStartupDialog: boolean): void
        getSelection(): any
        getStartupPath(): string
        getStatusBar(): any
        setStatusBar(statusBar: any): void
        getUsableHeight(): double
        getUsableWidth(): double
        isUserControl(): boolean
        isVisible(): boolean
        setVisible(visible: boolean): void
        getWindowState(): int
        setWindowState(windowState: int): void
        isShowWindowsInTaskbar(): boolean
        setShowWindowsInTaskbar(showWindowsInTaskbar: boolean): void
        getSmartTagRecognizers(): excel.SmartTagRecognizers
        getLanguageSettings(): office.LanguageSettings
        centimetersToPoints(centimeters: double): double
        checkSpelling(word: string, customDictionary: any, ignoreUppercase: any): boolean
        DDEExecute(channel: int, command: string): void
        DDETerminate(channel: int): void
        DDEInitiate(app: string, topic: string): int
        DDERequest(Channel: number, Item: string): string
        inchesToPoints(inches: double): double
        nextLetter(): excel.Workbook
        getDefaultWebOptions(): excel.DefaultWebOptions
        onTime(earliestTime: any, procedure: string, latestTime: any, schedule: any): void
        repeat(): void
        Run(macro: string): int
        Run(macro: string, args: long): int
        quit(): void
        removeApplicationListener(l: excel.event.ApplicationListener): void
        removeAllListeners(): void
        addApplicationListener(l: excel.event.ApplicationListener): void
        getNativeHandle(): int
        removeListener(book: excel.Workbook): void
        getStartInfo(): long
        fireEvent(type: int): void
        setGlobalInfo(globalInfo: int, handle: int): void
        setGlobalInfo(globalInfo: int): void
        setNativeHandle(handle: int): void
        getGlobalInfo(): int
        setStartInfo(startInfo: long): void
        getMenu(hMenu: long, type: int, index: int): long
        access$2(arg0: excel.Application): excel.event.MainAdapter
        access$3(arg0: excel.Application, arg1: excel.Workbook, arg2: excel.event.MainAdapter): void
        access$4(arg0: excel.Application): any
        createBinderHandle(mbook: any): void
        createPreviewPicForOle(): string
        getDefaultFilePath(): string
        getAllMSBarNames(): string[]
        getMWorkbook(): any
        setSaveFromStream(stream: boolean): void
        isSaveFromStream(): boolean
        isDisplayPasteOptions(): boolean
        setDisplayPasteOptions(displayPasteOptions: boolean): void
        getMeasurementUnit(): int
        setMeasurementUnit(measurementUnit: int): void
        getRange(cell1: any, cell2: any): excel.Range
        isEnableSound(): boolean
        setEnableSound(enableSound: boolean): void
        isMapPaperSize(): boolean
        setMapPaperSize(mapPaperSize: boolean): void
        getCursorMovement(): int
        setCursorMovement(cursorMovement: int): void
        setDefaultFilePath(defaultFilePath: string): void
        isAutoFormatAsYouTypeReplaceHyperlinks(): boolean
        setAutoFormatAsYouTypeReplaceHyperlinks(autoFormatAsYouTypeReplaceHyperlinks: boolean): void
        getCells(row: any, column: any): excel.Range
        getCells(): excel.Range
        getColumns(): excel.Range
        getRows(): excel.Range
        setUserControl(userControl: boolean): void
        setEnableEvents(enableEvents: boolean): void
        getWorksheets(): excel.Sheets
        getCharts(): excel.Sheets
        undo(): void
        isEnableEvents(): boolean
        calculate(): void
        getOperatingSystem(): string
        getActiveWorkbook(): excel.Workbook
        setOnWindow(onWindow: string): void
        getOnWindow(): string
        getWorkbooks(): excel.Workbooks
        getHwnd(): int
        getNames(): excel.Names
        isReady(): boolean
        getRTD(): excel.RTD
        evaluate(name: any): any
        findFile(): boolean
        goto1(reference: any, scroll: boolean): void
        goto1(reference: any, scroll: any): void
        inputBox(prompt: string, title: any, default1: any, left: any, top: any, helpFile: any, helpContextID: any, type: any): any
        onKey(key: string, procedure: any): void
        OnRepeat(text: string, procedure: string): void
        onUndo(text: string, procedure: string): void
        sendKeys(keys: any, wait: any): void
        wait1(time: int): boolean
        getMWorksheet(): any
        getMWorksheets(): any
        getMWorkbooks(): any
        getActiveCell(): excel.Range
        getActiveChart(): excel.Chart
        getActiveSheet(): any
        getAutoRecover(): excel.AutoRecover
        setCalculation(calculation: int): void
        getCalculation(): int
        getCanPlaySounds(): boolean
        setCutCopyMode(CutCopyMode: int): void
        getCutCopyMode(): int
        setDataEntryMode(dataEntryMode: int): void
        getDataEntryMode(): int
        isDisplayAlerts(): boolean
        isExtendList(): boolean
        setExtendList(extendList: boolean): void
        getFindFormat(): excel.CellFormat
        setFindFormat(format: excel.CellFormat): void
        isFixedDecimal(): boolean
        setFixedDecimal(fixedDecimal: boolean): void
        getHinstance(): int
        isInteractive(): boolean
        setInteractive(interactive: boolean): void
        isIteration(): boolean
        setIteration(iteration: boolean): void
        getLibraryPath(): string
        getMailSession(): any
        getMailSystem(): int
        getMaxChange(): double
        setMaxChange(maxChange: double): void
        getMaxIterations(): int
        getMemoryFree(): int
        getMemoryTotal(): int
        getMemoryUsed(): int
        setMaxIterations(maxIterations: int): void
        getNewWorkbook(): office.NewFile
        getODBCErrors(): excel.ODBCErrors
        getODBCTimeout(): int
        setODBCTimeout(timeout: int): void
        getOLEDBErrors(): excel.OLEDBErrors
        getProductCode(): string
        isRecordRelative(): boolean
        getReplaceFormat(): excel.CellFormat
        setReplaceFormat(format: excel.CellFormat): void
        isRollZoom(): boolean
        getSheets(): excel.Sheets
        isShowToolTips(): boolean
        setShowToolTips(showToolTips: boolean): void
        getSpeech(): excel.Speech
        getStandardFont(): string
        setStandardFont(standardFont: string): void
        getTemplatesPath(): string
        getThisCell(): excel.Range
        getThisWorkbook(): excel.Workbook
        getUsedObjects(): excel.UsedObjects
        getWatches(): excel.Watches
        isWindowsForPens(): boolean
        addCustomList(listArray: any, byRow: any): void
        calculateFull(): void
        checkAbort(keepAbort: any): void
        convertFormula(formula: any, fromReferenceStyle: int, toReferenceStyle: any, toAbsolute: any, relativeTo: any): any
        deleteCustomList(listNum: int): void
        doubleClick(): void
        getCustomListNum(listArray: any): int
        getOpenFilename(fileFilter: any, filterIndex: any, title: any, buttonText: any, multiSelect: any): any
        getPhonetic(text: any): string
        Intersect(arg1: excel.Range, arg2: excel.Range, args: excel.Range[]): excel.Range
        macroOptions(macro: any, description: any, hasMenu: any, menuText: any, hasShortcutKey: any, shortcutKey: any, category: any, statusBar: any, helpContextID: any, helpFile: any): void
        mailLogoff(): void
        mailLogon(name: any, password: any, downloadNewMail: any): void
        recordMacro(basicCode: any, xlmCode: any): void
        RegisterXLL(filename: string): boolean
        SaveWorkspace(filename: string): void
        setDefaultChart(formatName: string, gallery: any): void
        volatile1(volatile1: boolean): void
        addBookListener(book: excel.Workbook, l: excel.event.MainAdapter): void
        addSheetListener(book: excel.Workbook, applistener: excel.event.MainAdapter): void
        getMenuBars(): excel.MenuBars
        setAlertBeforeOverwriting(alertBeforeOverwriting: boolean): void
        setCalculationInterruptKey(calculationInterruptKey: int): void
        getCalculationInterruptKey(): int
        setDisplayClipboardWindow(DisplayClipboardWindow: boolean): void
        setDisplayCommentIndicator(displayCommentIndicator: int): void
        getDisplayCommentIndicator(): int
        setDisplayDocumentActionTaskPane(displayDocumentActionTaskPane: boolean): void
        isDisplayDocumentActionTaskPane(): boolean
        setDisplayFunctionToolTips(displayFunctionToolTips: boolean): void
        isDisplayFunctionToolTips(): boolean
        getMoveAfterReturnDirection(): int
        setMoveAfterReturnDirection(moveAfterReturnDirection: int): void
        getTransitionMenuKeyAction(): int
        setTransitionMenuKeyAction(transitionMenuKeyAction: int): void
        isAlertBeforeOverwriting(): boolean
        setAltStartupPath(altStartupPath: string): void
        getAltStartupPath(): string
        setAskToUpdateLinks(askToUpdateLinks: boolean): void
        isAskToUpdateLinks(): boolean
        setAutoPercentEntry(AutoPercentEntry: boolean): void
        isAutoPercentEntry(): boolean
        setCalculateBeforeSave(calculateBeforeSave: boolean): void
        isCalculateBeforeSave(): boolean
        getCalculationState(): int
        getCalculationVersion(): int
        getCanRecordSounds(): boolean
        setCellDragAndDrop(cellDragAndDrop: boolean): void
        isCellDragAndDrop(): boolean
        getClipboardFormats(index: any): any
        setCommandUnderlines(commandUnderlines: int): void
        getCommandUnderlines(): int
        setConstrainNumeric(constrainNumeric: boolean): void
        isConstrainNumeric(): boolean
        setControlCharacters(controlCharacters: boolean): void
        isControlCharacters(): boolean
        setCopyObjectsWithCells(copyObjectsWithCells: boolean): void
        isCopyObjectsWithCells(): boolean
        getCustomListCount(): int
        getDDEAppReturnCode(): int
        getDecimalSeparator(): string
        setDecimalSeparator(decimalSeparator: string): void
        getDefaultSheetDirection(): int
        setDefaultSheetDirection(defaultSheetDirection: int): void
        isDisplayClipboardWindow(): boolean
        setDisplayFormulaBar(displayFormulaBar: boolean): void
        isDisplayFormulaBar(): boolean
        setDisplayFullScreen(displayFullScreen: boolean): void
        isDisplayFullScreen(): boolean
        isDisplayInsertOptions(): boolean
        setDisplayInsertOptions(displayInsertOptions: boolean): void
        isDisplayNoteIndicator(): boolean
        setDisplayNoteIndicator(displayNoteIndicator: boolean): void
        setEditDirectlyInCell(editDirectlyInCell: boolean): void
        isEditDirectlyInCell(): boolean
        isEnableAnimations(): boolean
        setEnableAnimations(enableAnimations: boolean): void
        isEnableAutoComplete(): boolean
        setEnableAutoComplete(enableAutoComplete: boolean): void
        getErrorCheckingOptions(): excel.ErrorCheckingOptions
        getExcel4IntlMacroSheets(): excel.Sheets
        getExcel4MacroSheets(): excel.Sheets
        getFixedDecimalPlaces(): int
        setFixedDecimalPlaces(fixedDecimalPlaces: int): void
        getGenerateGetPivotData(): int
        setGenerateGetPivotData(generateGetPivotData: boolean): void
        isIgnoreRemoteRequests(): boolean
        setIgnoreRemoteRequests(ignoreRemoteRequests: boolean): void
        isMoveAfterReturn(): boolean
        setMoveAfterReturn(moveAfterReturn: boolean): void
        getNetworkTemplatesPath(): string
        getOrganizationName(): string
        getPreviousSelections(index: any): any
        isPromptForSummaryInfo(): boolean
        setPromptForSummaryInfo(promptForSummaryInfo: boolean): void
        getReferenceStyle(): int
        setReferenceStyle(referenceStyle: int): void
        getRegisteredFunctions(index1: any, index2: any): any
        getSheetsInNewWorkbook(): int
        setSheetsInNewWorkbook(sheetsInNewWorkbook: int): void
        isShowChartTipNames(): boolean
        setShowChartTipNames(showChartTipNames: boolean): void
        isShowChartTipValues(): boolean
        setShowChartTipValues(showChartTipValues: boolean): void
        getSpellingOptions(): excel.SpellingOptions
        getStandardFontSize(): double
        setStandardFontSize(standardFontSize: double): void
        getThousandsSeparator(): string
        setThousandsSeparator(thousandsSeparator: string): void
        getTransitionMenuKey(): string
        setTransitionMenuKey(transitionMenuKey: string): void
        isTransitionNavigKeys(): boolean
        setTransitionNavigKeys(transitionNavigKeys: boolean): void
        getUserLibraryPath(): string
        isUseSystemSeparators(): boolean
        setUseSystemSeparators(useSystemSeparators: boolean): void
        getWorksheetFunction(): excel.WorksheetFunction
        activateMicrosoftApp(index: int): void
        addChartAutoFormat(chart: excel.Chart, name: string, description: string): void
        calculateFullRebuild(): void
        DeleteChartAutoFormat(name: string): void
        displayXMLSourcePane(xmlmap: any): void
        executeExcel4Macro(string: string): any
        getCustomListContents(listNum: int): any
        getSaveAsFilename(initialFilename: any, fileFilter: any, filterIndex: any, title: any, buttonText: any): any
    }
    export interface Application {
        getName(): string
        getValue(): string
        close(saveChanges: int): void
        getPath(): string
        getPathSeparator(): string
        closeAll(): int
        help(helpFile: any, helpContextID: any): void
        union(arg1: excel.Range, arg2: excel.Range, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): excel.Range
        setCursor(cursor: int): void
        getCursor(): int
        getBuild(): int
        getWidth(): double
        getLeft(): double
        setLeft(left: double): void
        setTop(top: double): void
        getTop(): double
        setWidth(width: double): void
        DDEPoke(Channel: number, Item: string, Data: string): void
        getMainControl(): any
        getCommandBars(): office.CommandBars
        setActivePrinter(activePrinter: string): void
        getAutomationSecurity(): int
        setAutomationSecurity(AutomationSecurity: int): void
        isArbitraryXMLSupportAvailable(): boolean
        getActivePrinter(): string
        getActiveWindow(): excel.Window
        getWindows(): excel.Windows
        getAddIns(): excel.AddIns
        getAnswerWizard(): office.AnswerWizard
        getAssistant(): office.Assistant
        getAutoCorrect(): excel.AutoCorrect
        setCaption(caption: string): void
        getCaption(): string
        getCOMAddIns(): office.COMAddIns
        getOptions(): office.AbstractOptions
        getOptions(): excel.Options
        getDialogs(): excel.Dialogs
        setDisplayAlerts(displayAlerts: boolean): void
        getDefaultSaveFormat(): int
        setDefaultSaveFormat(defaultSaveFormat: int): void
        isDisplayRecentFiles(): boolean
        setDisplayRecentFiles(displayRecentFiles: boolean): void
        isDisplayScrollBars(): boolean
        setDisplayScrollBars(displayScrollBars: boolean): void
        isDisplayStatusBar(): boolean
        setDisplayStatusBar(displayStatusBar: boolean): void
        getEnableCancelKey(): int
        setEnableCancelKey(enableCancelKey: int): void
        getFeatureInstall(): int
        setFeatureInstall(featureInstall: int): void
        getFileConverters(index1: any, index2: any): any
        getFileDialog(fileDialogType: int): office.FileDialog
        getFileSearch(): office.FileSearch
        getHeight(): double
        setHeight(height: double): void
        getInternational(index: any): any
        isMouseAvailable(): boolean
        getRecentFiles(): excel.RecentFiles
        isScreenUpdating(): boolean
        isMathCoprocessorAvailable(): boolean
        setScreenUpdating(screenUpdating: boolean): void
        isShowStartupDialog(): boolean
        setShowStartupDialog(showStartupDialog: boolean): void
        getSelection(): any
        getStartupPath(): string
        getStatusBar(): any
        setStatusBar(statusBar: any): void
        getUsableHeight(): double
        getUsableWidth(): double
        isUserControl(): boolean
        isVisible(): boolean
        setVisible(visible: boolean): void
        getWindowState(): int
        setWindowState(windowState: int): void
        isShowWindowsInTaskbar(): boolean
        setShowWindowsInTaskbar(showWindowsInTaskbar: boolean): void
        getSmartTagRecognizers(): excel.SmartTagRecognizers
        getLanguageSettings(): office.LanguageSettings
        centimetersToPoints(centimeters: double): double
        checkSpelling(word: string, customDictionary: any, ignoreUppercase: any): boolean
        DDEExecute(channel: int, command: string): void
        DDETerminate(channel: int): void
        DDEInitiate(app: string, topic: string): int
        DDERequest(Channel: number, Item: string): string
        inchesToPoints(inches: double): double
        nextLetter(): excel.Workbook
        getDefaultWebOptions(): excel.DefaultWebOptions
        onTime(earliestTime: any, procedure: string, latestTime: any, schedule: any): void
        repeat(): void
        Run(macro: string): int
        Run(macro: string, args: long): int
        quit(): void
        removeApplicationListener(l: excel.event.ApplicationListener): void
        removeAllListeners(): void
        addApplicationListener(l: excel.event.ApplicationListener): void
        getNativeHandle(): int
        removeListener(book: excel.Workbook): void
        getStartInfo(): long
        fireEvent(type: int): void
        setGlobalInfo(globalInfo: int, handle: int): void
        setGlobalInfo(globalInfo: int): void
        setNativeHandle(handle: int): void
        getGlobalInfo(): int
        setStartInfo(startInfo: long): void
        getMenu(hMenu: long, type: int, index: int): long
        access$2(arg0: excel.Application): excel.event.MainAdapter
        access$3(arg0: excel.Application, arg1: excel.Workbook, arg2: excel.event.MainAdapter): void
        access$4(arg0: excel.Application): any
        createBinderHandle(mbook: any): void
        createPreviewPicForOle(): string
        getDefaultFilePath(): string
        getAllMSBarNames(): string[]
        getMWorkbook(): any
        setSaveFromStream(stream: boolean): void
        isSaveFromStream(): boolean
        isDisplayPasteOptions(): boolean
        setDisplayPasteOptions(displayPasteOptions: boolean): void
        getMeasurementUnit(): int
        setMeasurementUnit(measurementUnit: int): void
        getRange(cell1: any, cell2: any): excel.Range
        isEnableSound(): boolean
        setEnableSound(enableSound: boolean): void
        isMapPaperSize(): boolean
        setMapPaperSize(mapPaperSize: boolean): void
        getCursorMovement(): int
        setCursorMovement(cursorMovement: int): void
        setDefaultFilePath(defaultFilePath: string): void
        isAutoFormatAsYouTypeReplaceHyperlinks(): boolean
        setAutoFormatAsYouTypeReplaceHyperlinks(autoFormatAsYouTypeReplaceHyperlinks: boolean): void
        getCells(row: any, column: any): excel.Range
        getCells(): excel.Range
        getColumns(): excel.Range
        getRows(): excel.Range
        setUserControl(userControl: boolean): void
        setEnableEvents(enableEvents: boolean): void
        getWorksheets(): excel.Sheets
        getCharts(): excel.Sheets
        undo(): void
        isEnableEvents(): boolean
        calculate(): void
        getOperatingSystem(): string
        getActiveWorkbook(): excel.Workbook
        setOnWindow(onWindow: string): void
        getOnWindow(): string
        getWorkbooks(): excel.Workbooks
        getHwnd(): int
        getNames(): excel.Names
        isReady(): boolean
        getRTD(): excel.RTD
        evaluate(name: any): any
        findFile(): boolean
        goto1(reference: any, scroll: boolean): void
        goto1(reference: any, scroll: any): void
        inputBox(prompt: string, title: any, default1: any, left: any, top: any, helpFile: any, helpContextID: any, type: any): any
        onKey(key: string, procedure: any): void
        OnRepeat(text: string, procedure: string): void
        onUndo(text: string, procedure: string): void
        sendKeys(keys: any, wait: any): void
        wait1(time: int): boolean
        getMWorksheet(): any
        getMWorksheets(): any
        getMWorkbooks(): any
        getActiveCell(): excel.Range
        getActiveChart(): excel.Chart
        getActiveSheet(): any
        getAutoRecover(): excel.AutoRecover
        setCalculation(calculation: int): void
        getCalculation(): int
        getCanPlaySounds(): boolean
        setCutCopyMode(CutCopyMode: int): void
        getCutCopyMode(): int
        setDataEntryMode(dataEntryMode: int): void
        getDataEntryMode(): int
        isDisplayAlerts(): boolean
        isExtendList(): boolean
        setExtendList(extendList: boolean): void
        getFindFormat(): excel.CellFormat
        setFindFormat(format: excel.CellFormat): void
        isFixedDecimal(): boolean
        setFixedDecimal(fixedDecimal: boolean): void
        getHinstance(): int
        isInteractive(): boolean
        setInteractive(interactive: boolean): void
        isIteration(): boolean
        setIteration(iteration: boolean): void
        getLibraryPath(): string
        getMailSession(): any
        getMailSystem(): int
        getMaxChange(): double
        setMaxChange(maxChange: double): void
        getMaxIterations(): int
        getMemoryFree(): int
        getMemoryTotal(): int
        getMemoryUsed(): int
        setMaxIterations(maxIterations: int): void
        getNewWorkbook(): office.NewFile
        getODBCErrors(): excel.ODBCErrors
        getODBCTimeout(): int
        setODBCTimeout(timeout: int): void
        getOLEDBErrors(): excel.OLEDBErrors
        getProductCode(): string
        isRecordRelative(): boolean
        getReplaceFormat(): excel.CellFormat
        setReplaceFormat(format: excel.CellFormat): void
        isRollZoom(): boolean
        getSheets(): excel.Sheets
        isShowToolTips(): boolean
        setShowToolTips(showToolTips: boolean): void
        getSpeech(): excel.Speech
        getStandardFont(): string
        setStandardFont(standardFont: string): void
        getTemplatesPath(): string
        getThisCell(): excel.Range
        getThisWorkbook(): excel.Workbook
        getUsedObjects(): excel.UsedObjects
        getWatches(): excel.Watches
        isWindowsForPens(): boolean
        addCustomList(listArray: any, byRow: any): void
        calculateFull(): void
        checkAbort(keepAbort: any): void
        convertFormula(formula: any, fromReferenceStyle: int, toReferenceStyle: any, toAbsolute: any, relativeTo: any): any
        deleteCustomList(listNum: int): void
        doubleClick(): void
        getCustomListNum(listArray: any): int
        getOpenFilename(fileFilter: any, filterIndex: any, title: any, buttonText: any, multiSelect: any): any
        getPhonetic(text: any): string
        Intersect(arg1: excel.Range, arg2: excel.Range, args: excel.Range[]): excel.Range
        macroOptions(macro: any, description: any, hasMenu: any, menuText: any, hasShortcutKey: any, shortcutKey: any, category: any, statusBar: any, helpContextID: any, helpFile: any): void
        mailLogoff(): void
        mailLogon(name: any, password: any, downloadNewMail: any): void
        recordMacro(basicCode: any, xlmCode: any): void
        RegisterXLL(filename: string): boolean
        SaveWorkspace(filename: string): void
        setDefaultChart(formatName: string, gallery: any): void
        volatile1(volatile1: boolean): void
        addBookListener(book: excel.Workbook, l: excel.event.MainAdapter): void
        addSheetListener(book: excel.Workbook, applistener: excel.event.MainAdapter): void
        getMenuBars(): excel.MenuBars
        setAlertBeforeOverwriting(alertBeforeOverwriting: boolean): void
        setCalculationInterruptKey(calculationInterruptKey: int): void
        getCalculationInterruptKey(): int
        setDisplayClipboardWindow(DisplayClipboardWindow: boolean): void
        setDisplayCommentIndicator(displayCommentIndicator: int): void
        getDisplayCommentIndicator(): int
        setDisplayDocumentActionTaskPane(displayDocumentActionTaskPane: boolean): void
        isDisplayDocumentActionTaskPane(): boolean
        setDisplayFunctionToolTips(displayFunctionToolTips: boolean): void
        isDisplayFunctionToolTips(): boolean
        getMoveAfterReturnDirection(): int
        setMoveAfterReturnDirection(moveAfterReturnDirection: int): void
        getTransitionMenuKeyAction(): int
        setTransitionMenuKeyAction(transitionMenuKeyAction: int): void
        isAlertBeforeOverwriting(): boolean
        setAltStartupPath(altStartupPath: string): void
        getAltStartupPath(): string
        setAskToUpdateLinks(askToUpdateLinks: boolean): void
        isAskToUpdateLinks(): boolean
        setAutoPercentEntry(AutoPercentEntry: boolean): void
        isAutoPercentEntry(): boolean
        setCalculateBeforeSave(calculateBeforeSave: boolean): void
        isCalculateBeforeSave(): boolean
        getCalculationState(): int
        getCalculationVersion(): int
        getCanRecordSounds(): boolean
        setCellDragAndDrop(cellDragAndDrop: boolean): void
        isCellDragAndDrop(): boolean
        getClipboardFormats(index: any): any
        setCommandUnderlines(commandUnderlines: int): void
        getCommandUnderlines(): int
        setConstrainNumeric(constrainNumeric: boolean): void
        isConstrainNumeric(): boolean
        setControlCharacters(controlCharacters: boolean): void
        isControlCharacters(): boolean
        setCopyObjectsWithCells(copyObjectsWithCells: boolean): void
        isCopyObjectsWithCells(): boolean
        getCustomListCount(): int
        getDDEAppReturnCode(): int
        getDecimalSeparator(): string
        setDecimalSeparator(decimalSeparator: string): void
        getDefaultSheetDirection(): int
        setDefaultSheetDirection(defaultSheetDirection: int): void
        isDisplayClipboardWindow(): boolean
        setDisplayFormulaBar(displayFormulaBar: boolean): void
        isDisplayFormulaBar(): boolean
        setDisplayFullScreen(displayFullScreen: boolean): void
        isDisplayFullScreen(): boolean
        isDisplayInsertOptions(): boolean
        setDisplayInsertOptions(displayInsertOptions: boolean): void
        isDisplayNoteIndicator(): boolean
        setDisplayNoteIndicator(displayNoteIndicator: boolean): void
        setEditDirectlyInCell(editDirectlyInCell: boolean): void
        isEditDirectlyInCell(): boolean
        isEnableAnimations(): boolean
        setEnableAnimations(enableAnimations: boolean): void
        isEnableAutoComplete(): boolean
        setEnableAutoComplete(enableAutoComplete: boolean): void
        getErrorCheckingOptions(): excel.ErrorCheckingOptions
        getExcel4IntlMacroSheets(): excel.Sheets
        getExcel4MacroSheets(): excel.Sheets
        getFixedDecimalPlaces(): int
        setFixedDecimalPlaces(fixedDecimalPlaces: int): void
        getGenerateGetPivotData(): int
        setGenerateGetPivotData(generateGetPivotData: boolean): void
        isIgnoreRemoteRequests(): boolean
        setIgnoreRemoteRequests(ignoreRemoteRequests: boolean): void
        isMoveAfterReturn(): boolean
        setMoveAfterReturn(moveAfterReturn: boolean): void
        getNetworkTemplatesPath(): string
        getOrganizationName(): string
        getPreviousSelections(index: any): any
        isPromptForSummaryInfo(): boolean
        setPromptForSummaryInfo(promptForSummaryInfo: boolean): void
        getReferenceStyle(): int
        setReferenceStyle(referenceStyle: int): void
        getRegisteredFunctions(index1: any, index2: any): any
        getSheetsInNewWorkbook(): int
        setSheetsInNewWorkbook(sheetsInNewWorkbook: int): void
        isShowChartTipNames(): boolean
        setShowChartTipNames(showChartTipNames: boolean): void
        isShowChartTipValues(): boolean
        setShowChartTipValues(showChartTipValues: boolean): void
        getSpellingOptions(): excel.SpellingOptions
        getStandardFontSize(): double
        setStandardFontSize(standardFontSize: double): void
        getThousandsSeparator(): string
        setThousandsSeparator(thousandsSeparator: string): void
        getTransitionMenuKey(): string
        setTransitionMenuKey(transitionMenuKey: string): void
        isTransitionNavigKeys(): boolean
        setTransitionNavigKeys(transitionNavigKeys: boolean): void
        getUserLibraryPath(): string
        isUseSystemSeparators(): boolean
        setUseSystemSeparators(useSystemSeparators: boolean): void
        getWorksheetFunction(): excel.WorksheetFunction
        activateMicrosoftApp(index: int): void
        addChartAutoFormat(chart: excel.Chart, name: string, description: string): void
        calculateFullRebuild(): void
        DeleteChartAutoFormat(name: string): void
        displayXMLSourcePane(xmlmap: any): void
        executeExcel4Macro(string: string): any
        getCustomListContents(listNum: int): any
        getSaveAsFilename(initialFilename: any, fileFilter: any, filterIndex: any, title: any, buttonText: any): any
    }
    export interface Areas {
        item(index: int): excel.Range
        getCount(): int
    }
    export interface Areas {
        item(index: int): excel.Range
        getCount(): int
    }
    export interface AutoCorrect {
        setReplaceText(replaceText: boolean): void
        isReplaceText(): boolean
        setDisplayAutoCorrectOptions(displayAutoCorrectOptions: boolean): void
        isDisplayAutoCorrectOptions(): boolean
        setCorrectCapsLock(correctCapsLock: boolean): void
        isCorrectCapsLock(): boolean
        isAutoExpandListRange(): boolean
        setAutoExpandListRange(autoExpandListRange: boolean): void
        isCapitalizeNamesOfDays(): boolean
        setCapitalizeNamesOfDays(capitalizeNamesOfDays: boolean): void
        isCorrectSentenceCap(): boolean
        setCorrectSentenceCap(correctSentenceCap: boolean): void
        isTwoInitialCapitals(): boolean
        setTwoInitialCapitals(twoInitialCapitals: boolean): void
        getReplacementList(): any
        getReplacementList(index: any): any
        deleteReplacement(what: string): any
        addReplacement(what: string, replacement: string): any
    }
    export interface AutoCorrect {
        setReplaceText(replaceText: boolean): void
        isReplaceText(): boolean
        setDisplayAutoCorrectOptions(displayAutoCorrectOptions: boolean): void
        isDisplayAutoCorrectOptions(): boolean
        setCorrectCapsLock(correctCapsLock: boolean): void
        isCorrectCapsLock(): boolean
        isAutoExpandListRange(): boolean
        setAutoExpandListRange(autoExpandListRange: boolean): void
        isCapitalizeNamesOfDays(): boolean
        setCapitalizeNamesOfDays(capitalizeNamesOfDays: boolean): void
        isCorrectSentenceCap(): boolean
        setCorrectSentenceCap(correctSentenceCap: boolean): void
        isTwoInitialCapitals(): boolean
        setTwoInitialCapitals(twoInitialCapitals: boolean): void
        getReplacementList(): any
        getReplacementList(index: any): any
        deleteReplacement(what: string): any
        addReplacement(what: string, replacement: string): any
    }
    export interface AutoFilter {
        getRange(): excel.Range
        getFilters(): excel.Filters
        ApplyFilter(): void
        filterMode(): boolean
        showAllData(): void
    }
    export interface AutoFilter {
        getRange(): excel.Range
        getFilters(): excel.Filters
        ApplyFilter(): void
        filterMode(): boolean
        showAllData(): void
    }
    export interface AutoRecover {
        getPath(): string
        setTime(time: int): void
        getTime(): int
        setEnabled(enabled: boolean): void
        isEnabled(): boolean
        setPath(path: string): void
    }
    export interface AutoRecover {
        getPath(): string
        setTime(time: int): void
        getTime(): int
        setEnabled(enabled: boolean): void
        isEnabled(): boolean
        setPath(path: string): void
    }
    export interface Axes {
        item(type: int, axisGroup: int): excel.Axis
        getCount(): int
    }
    export interface Axes {
        item(type: int, axisGroup: int): excel.Axis
        getCount(): int
    }
    export interface Axis {
        delete(): any
        getType(): int
        getBorder(): excel.Border
        getWidth(): double
        getLeft(): double
        getTop(): double
        getHeight(): double
        select(): any
        getDisplayUnitLabel(): excel.DisplayUnitLabel
        isAxisBetweenCategories(): boolean
        setAxisBetweenCategories(axisBetweenCategories: boolean): void
        setBaseUnitIsAuto(baseUnitIsAuto: boolean): void
        getDisplayUnitCustom(): double
        setDisplayUnitCustom(displayUnitCustom: double): void
        isHasDisplayUnitLabel(): boolean
        setHasDisplayUnitLabel(hasDisplayUnitLabel: boolean): void
        isHasMajorGridlines(): boolean
        setHasMajorGridlines(hasMajorGridlines: boolean): void
        isHasMinorGridlines(): boolean
        setHasMinorGridlines(hasMinorGridlines: boolean): void
        getMajorGridlines(): excel.Gridlines
        isMajorUnitIsAuto(): boolean
        setMajorUnitIsAuto(majorUnitIsAuto: boolean): void
        getMajorUnitScale(): int
        setMajorUnitScale(majorUnitScale: int): void
        isMaximumScaleIsAuto(): boolean
        setMaximumScaleIsAuto(maximumScaleIsAuto: boolean): void
        isMinimumScaleIsAuto(): boolean
        setMinimumScaleIsAuto(minimumScaleIsAuto: boolean): void
        getMinorGridlines(): excel.Gridlines
        isMinorUnitIsAuto(): boolean
        setMinorUnitIsAuto(minorUnitIsAuto: boolean): void
        getMinorUnitScale(): int
        setMinorUnitScale(minorUnitScale: int): void
        isReversePlotOrder(): boolean
        setReversePlotOrder(reversePlotOrder: boolean): void
        getTickLabelPosition(): int
        setTickLabelPosition(tickLabelPosition: int): void
        getTickLabelSpacing(): int
        setTickLabelSpacing(tickLabelSpacing: int): void
        getTickMarkSpacing(): int
        setTickMarkSpacing(tickMarkSpacing: int): void
        getAxisGroup(): int
        getAxisTitle(): excel.AxisTitle
        getBaseUnit(): int
        setBaseUnit(baseUnit: int): void
        isBaseUnitIsAuto(): boolean
        getCategoryType(): int
        setCategoryType(categoryType: int): void
        getCrosses(): int
        setCrosses(crosses: int): void
        getCrossesAt(): double
        setCrossesAt(crossesAt: double): void
        getDisplayUnit(): int
        setDisplayUnit(displayUnit: int): void
        getCategoryNames(): any
        setCategoryNames(categoryNames: any): void
        isHasTitle(): boolean
        setHasTitle(hasTitle: boolean): void
        getMajorTickMark(): int
        setMajorTickMark(majorTickMark: int): void
        getMajorUnit(): double
        setMajorUnit(majorUnit: double): void
        getMaximumScale(): double
        setMaximumScale(maximumScale: double): void
        getMinimumScale(): double
        setMinimumScale(minimumScale: double): void
        getTickLabels(): excel.TickLabels
        getMinorTickMark(): int
        setMinorTickMark(minorTickMark: int): void
        getMinorUnit(): double
        setMinorUnit(minorUnit: double): void
        getScaleType(): int
        setScaleType(scaleType: int): void
    }
    export interface Axis {
        delete(): any
        getType(): int
        getBorder(): excel.Border
        getWidth(): double
        getLeft(): double
        getTop(): double
        getHeight(): double
        select(): any
        getDisplayUnitLabel(): excel.DisplayUnitLabel
        isAxisBetweenCategories(): boolean
        setAxisBetweenCategories(axisBetweenCategories: boolean): void
        setBaseUnitIsAuto(baseUnitIsAuto: boolean): void
        getDisplayUnitCustom(): double
        setDisplayUnitCustom(displayUnitCustom: double): void
        isHasDisplayUnitLabel(): boolean
        setHasDisplayUnitLabel(hasDisplayUnitLabel: boolean): void
        isHasMajorGridlines(): boolean
        setHasMajorGridlines(hasMajorGridlines: boolean): void
        isHasMinorGridlines(): boolean
        setHasMinorGridlines(hasMinorGridlines: boolean): void
        getMajorGridlines(): excel.Gridlines
        isMajorUnitIsAuto(): boolean
        setMajorUnitIsAuto(majorUnitIsAuto: boolean): void
        getMajorUnitScale(): int
        setMajorUnitScale(majorUnitScale: int): void
        isMaximumScaleIsAuto(): boolean
        setMaximumScaleIsAuto(maximumScaleIsAuto: boolean): void
        isMinimumScaleIsAuto(): boolean
        setMinimumScaleIsAuto(minimumScaleIsAuto: boolean): void
        getMinorGridlines(): excel.Gridlines
        isMinorUnitIsAuto(): boolean
        setMinorUnitIsAuto(minorUnitIsAuto: boolean): void
        getMinorUnitScale(): int
        setMinorUnitScale(minorUnitScale: int): void
        isReversePlotOrder(): boolean
        setReversePlotOrder(reversePlotOrder: boolean): void
        getTickLabelPosition(): int
        setTickLabelPosition(tickLabelPosition: int): void
        getTickLabelSpacing(): int
        setTickLabelSpacing(tickLabelSpacing: int): void
        getTickMarkSpacing(): int
        setTickMarkSpacing(tickMarkSpacing: int): void
        getAxisGroup(): int
        getAxisTitle(): excel.AxisTitle
        getBaseUnit(): int
        setBaseUnit(baseUnit: int): void
        isBaseUnitIsAuto(): boolean
        getCategoryType(): int
        setCategoryType(categoryType: int): void
        getCrosses(): int
        setCrosses(crosses: int): void
        getCrossesAt(): double
        setCrossesAt(crossesAt: double): void
        getDisplayUnit(): int
        setDisplayUnit(displayUnit: int): void
        getCategoryNames(): any
        setCategoryNames(categoryNames: any): void
        isHasTitle(): boolean
        setHasTitle(hasTitle: boolean): void
        getMajorTickMark(): int
        setMajorTickMark(majorTickMark: int): void
        getMajorUnit(): double
        setMajorUnit(majorUnit: double): void
        getMaximumScale(): double
        setMaximumScale(maximumScale: double): void
        getMinimumScale(): double
        setMinimumScale(minimumScale: double): void
        getTickLabels(): excel.TickLabels
        getMinorTickMark(): int
        setMinorTickMark(minorTickMark: int): void
        getMinorUnit(): double
        setMinorUnit(minorUnit: double): void
        getScaleType(): int
        setScaleType(scaleType: int): void
    }
    export interface AxisTitle {
        getName(): string
        delete(): any
        getBorder(): excel.Border
        getFont(): excel.Font
        getOrientation(): any
        getLeft(): double
        setLeft(left: double): void
        setTop(top: double): void
        getTop(): double
        setCaption(caption: string): void
        getCaption(): string
        getText(): string
        setText(text: string): void
        select(): any
        isShadow(): boolean
        setShadow(shadow: boolean): void
        setVerticalAlignment(verticalAlignment: any): void
        getVerticalAlignment(): any
        getFill(): excel.ChartFillFormat
        setOrientation(orientation: any): void
        getCharacters(start: any, length: any): excel.Characters
        getReadingOrder(): int
        setReadingOrder(readingOrder: int): void
        getHorizontalAlignment(): any
        setHorizontalAlignment(horizontalAlignment: any): void
        getInterior(): excel.Interior
        isAutoScaleFont(): boolean
        setAutoScaleFont(autoScaleFont: boolean): void
    }
    export interface AxisTitle {
        getName(): string
        delete(): any
        getBorder(): excel.Border
        getFont(): excel.Font
        getOrientation(): any
        getLeft(): double
        setLeft(left: double): void
        setTop(top: double): void
        getTop(): double
        setCaption(caption: string): void
        getCaption(): string
        getText(): string
        setText(text: string): void
        select(): any
        isShadow(): boolean
        setShadow(shadow: boolean): void
        setVerticalAlignment(verticalAlignment: any): void
        getVerticalAlignment(): any
        getFill(): excel.ChartFillFormat
        setOrientation(orientation: any): void
        getCharacters(start: any, length: any): excel.Characters
        getReadingOrder(): int
        setReadingOrder(readingOrder: int): void
        getHorizontalAlignment(): any
        setHorizontalAlignment(horizontalAlignment: any): void
        getInterior(): excel.Interior
        isAutoScaleFont(): boolean
        setAutoScaleFont(autoScaleFont: boolean): void
    }
    export interface Border {
        setColor(color: any, borderAttribute: any): void
        setColor(obj: any): void
        getColor(): any
        getColor(borderAttribute: any): any
        setColorIndex(index: any): void
        getColorIndex(): any
        setLineStyle(style: int, width: float, borderAttribute: any): void
        setLineStyle(obj: any): void
        getLineStyle(borderAttribute: any): int
        getLineStyle(): any
        getWeight(): any
        setWeight(obj: any): void
        setWeight(weight: int, borderAttribute: any): void
    }
    export interface Border {
        setColor(color: any, borderAttribute: any): void
        setColor(obj: any): void
        getColor(): any
        getColor(borderAttribute: any): any
        setColorIndex(index: any): void
        getColorIndex(): any
        setLineStyle(style: int, width: float, borderAttribute: any): void
        setLineStyle(obj: any): void
        getLineStyle(borderAttribute: any): int
        getLineStyle(): any
        getWeight(): any
        setWeight(obj: any): void
        setWeight(weight: int, borderAttribute: any): void
    }
    export interface Borders {
        getValue(): any
        setValue(value: any): void
        setColor(color: any, borderAttribute: any): void
        setColor(rgb: any): void
        item(borderIndex: any): excel.Border
        getCount(): int
        getItem(borderIndex: int): excel.Border
        getColor(): any
        getColor(borderAttribute: any): any
        setColorIndex(index: any): void
        getColorIndex(): any
        setLineStyle(lineStyle: any): void
        setLineStyle(style: int, width: float, borderAttribute: any): void
        getLineStyle(borderAttribute: any): int
        getLineStyle(): any
        getWeight(): any
        setWeight(weight: int, borderAttribute: any): boolean
        setWeight(weight: any): void
        setBorderAttr(borderAttr: any): void
    }
    export interface Borders {
        getValue(): any
        setValue(value: any): void
        setColor(color: any, borderAttribute: any): void
        setColor(rgb: any): void
        item(borderIndex: any): excel.Border
        getCount(): int
        getItem(borderIndex: int): excel.Border
        getColor(): any
        getColor(borderAttribute: any): any
        setColorIndex(index: any): void
        getColorIndex(): any
        setLineStyle(lineStyle: any): void
        setLineStyle(style: int, width: float, borderAttribute: any): void
        getLineStyle(borderAttribute: any): int
        getLineStyle(): any
        getWeight(): any
        setWeight(weight: int, borderAttribute: any): boolean
        setWeight(weight: any): void
        setBorderAttr(borderAttr: any): void
    }
    export interface CalculatedFields {
        add(name: string, formula: string, useStandardFormula: any): excel.PivotField
        item(index: any): excel.PivotField
        getCount(): int
    }
    export interface CalculatedFields {
        add(name: string, formula: string, useStandardFormula: any): excel.PivotField
        item(index: any): excel.PivotField
        getCount(): int
    }
    export interface CalculatedItems {
        add(name: string, formula: string, useStandardFormula: any): excel.PivotItem
        item(index: any): excel.PivotItem
        getCount(): int
    }
    export interface CalculatedItems {
        add(name: string, formula: string, useStandardFormula: any): excel.PivotItem
        item(index: any): excel.PivotItem
        getCount(): int
    }
    export interface CalculatedMember {
        getName(): string
        delete(): void
        getType(): int
        isValid(): boolean
        getSourceName(): string
        getFormula(): string
        getSolveOrder(): int
    }
    export interface CalculatedMember {
        getName(): string
        delete(): void
        getType(): int
        isValid(): boolean
        getSourceName(): string
        getFormula(): string
        getSolveOrder(): int
    }
    export interface CalculatedMembers {
        add(name: string, formula: string, solveOrder: int, type: int): excel.CalculatedMember
        item(index: int): excel.CalculatedMember
        getCount(): int
    }
    export interface CalculatedMembers {
        add(name: string, formula: string, solveOrder: int, type: int): excel.CalculatedMember
        item(index: int): excel.CalculatedMember
        getCount(): int
    }
    export interface CalloutFormat {
        getLength(): double
        getType(): int
        setBorder(border: int): void
        setType(type: int): void
        isAccent(): boolean
        setAngle(angle: int): void
        getAngle(): int
        isBorder(): boolean
        getDrop(): double
        setGap(gap: double): void
        getGap(): float
        presetDrop(DropType: int): void
        setAccent(accent: int): void
        setAutoAttach(auto: int): void
        isAutoAttach(): boolean
        isAutoLength(): boolean
        automaticLength(): void
        customDrop(drop: double): void
        customLength(length: double): void
        getDropType(): int
    }
    export interface CalloutFormat {
        getLength(): double
        getType(): int
        setBorder(border: int): void
        setType(type: int): void
        isAccent(): boolean
        setAngle(angle: int): void
        getAngle(): int
        isBorder(): boolean
        getDrop(): double
        setGap(gap: double): void
        getGap(): float
        presetDrop(DropType: int): void
        setAccent(accent: int): void
        setAutoAttach(auto: int): void
        isAutoAttach(): boolean
        isAutoLength(): boolean
        automaticLength(): void
        customDrop(drop: double): void
        customLength(length: double): void
        getDropType(): int
    }
    export interface CellFormat {
        clear(): void
        isLocked(): any
        getFont(): excel.Font
        getOrientation(): any
        setFont(font: excel.Font): void
        setBorders(borders: excel.Borders): void
        getBorders(): excel.Borders
        setVerticalAlignment(alignment: any): void
        getVerticalAlignment(): any
        setOrientation(orientation: any): void
        getNumberFormat(): any
        setNumberFormat(format: any): void
        setLocked(locked: any): void
        getHorizontalAlignment(): any
        setHorizontalAlignment(alignment: any): void
        getInterior(): excel.Interior
        setWrapText(wrap: any): void
        isAddIndent(): any
        setAddIndent(addIndent: any): void
        isFormulaHidden(): any
        setFormulaHidden(formulaHidden: any): void
        setIndentLevel(level: any): void
        getIndentLevel(): any
        setInterior(interior: excel.Interior): void
        setMergeCells(merge: any): void
        isMergeCells(): any
        setShrinkToFit(fit: any): void
        isShrinkToFit(): any
        isWrapText(): any
        setNumberFormatLocal(format: any): void
        getNumberFormatLocal(): any
    }
    export interface CellFormat {
        clear(): void
        isLocked(): any
        getFont(): excel.Font
        getOrientation(): any
        setFont(font: excel.Font): void
        setBorders(borders: excel.Borders): void
        getBorders(): excel.Borders
        setVerticalAlignment(alignment: any): void
        getVerticalAlignment(): any
        setOrientation(orientation: any): void
        getNumberFormat(): any
        setNumberFormat(format: any): void
        setLocked(locked: any): void
        getHorizontalAlignment(): any
        setHorizontalAlignment(alignment: any): void
        getInterior(): excel.Interior
        setWrapText(wrap: any): void
        isAddIndent(): any
        setAddIndent(addIndent: any): void
        isFormulaHidden(): any
        setFormulaHidden(formulaHidden: any): void
        setIndentLevel(level: any): void
        getIndentLevel(): any
        setInterior(interior: excel.Interior): void
        setMergeCells(merge: any): void
        isMergeCells(): any
        setShrinkToFit(fit: any): void
        isShrinkToFit(): any
        isWrapText(): any
        setNumberFormatLocal(format: any): void
        getNumberFormatLocal(): any
    }
    export interface Characters {
        delete(): any
        insert(string: string): any
        getFont(): excel.Font
        getCount(): int
        getCaption(): string
        getText(): string
        setText(text: string): void
        getMCharacters(): any
        getPhoneticCharacters(): string
        setPhoneticCharacters(phonetic: string): void
    }
    export interface Characters {
        delete(): any
        insert(string: string): any
        getFont(): excel.Font
        getCount(): int
        getCaption(): string
        getText(): string
        setText(text: string): void
        getMCharacters(): any
        getPhoneticCharacters(): string
        setPhoneticCharacters(phonetic: string): void
    }
    export interface Chart {
        getName(): string
        delete(): void
        setName(name: string): void
        copy(before: any, after: any): void
        location(where: int, nameObj: any): excel.Chart
        getIndex(): int
        refresh(): void
        activate(): void
        printPreview(enableChanges: any): void
        setVisible(visible: int): void
        checkSpelling(customDictionary: any, ignoreUppercase: any, alwaysSuggest: any, spellLang: any): void
        move(before: any, after: any): void
        printOut(from: any, to: any, copies: any, preview: any, activePrinter: any, printToFile: any, collate: any, prToFileName: any): void
        saveAs(filename: string, fileFormat: any, password: any, writeResPassword: any, readOnlyRecommended: any, createBackup: any, addToMru: any, textCodepage: any, textVisualLayout: any, local: any): void
        paste(type: any): void
        getPrevious(): any
        select(replace: any): void
        getPageSetup(): excel.PageSetup
        getNext(): any
        getVisible(): int
        getMailEnvelope(): int
        getScripts(): office.Scripts
        getShapes(): excel.Shapes
        unprotect(password: any): void
        exportAsFixedFormat(type: int, filename: any, quality: any, includeDocProperties: any, ignorePrintAreas: any, from: any, to: any, openAfterPublish: any, fixedFormatExtClassPtr: any): void
        protect(password: any, drawingObjects: any, contents: any, scenarios: any, userInterfaceOnly: any): void
        getCodeName(): string
        getHyperlinks(): excel.Hyperlinks
        createPublisher(edition: any, appearance: int, size: int, containsPICT: any, containsBIFF: any, containsRTF: any, containsVALU: any): void
        getRotation(): any
        setRotation(rotation: any): void
        setPerspective(perspective: int): void
        getPerspective(): int
        evaluate(name: any): any
        axes(type: any, axisGroup: int): any
        getAxes(): excel.Axes
        getFloor(): excel.Floor
        getTab(): excel.Tab
        getWalls(): excel.Walls
        deselect(): void
        export(filename: string, filterName: any, interactive: any): boolean
        xYGroups(index: int): any
        copyPicture(appearance: int, format: int, size: int): void
        columnGroups(index: int): any
        doughnutGroups(index: int): any
        getChartElement(x: int, y: int, ElementID: int, Arg1: int, Arg2: int): void
        OLEObjects(index: any): any
        pieGroups(index: int): any
        radarGroups(index: int): any
        setSourceData(source: excel.Range, plotBy: any): void
        surfaceGroups(index: int): any
        transTypeToEIO(xlsChartType: int): number[]
        transTypeToExcel(type: int, subType: int): int
        getMChart(): any
        isHasTitle(): boolean
        setHasTitle(hasTitle: boolean): void
        chartGroups(index: any): any
        seriesCollection(index: any): any
        getArea3DGroup(): excel.ChartGroup
        getChartType(): int
        setAutoScaling(autoScaling: boolean): void
        isAutoScaling(): boolean
        getBar3DGroup(): excel.ChartGroup
        getBarShape(): int
        setBarShape(barShape: int): void
        getChartArea(): excel.ChartArea
        getChartTitle(): excel.ChartTitle
        setChartStyle(chartStyle: any): void
        getChartStyle(): any
        setChartType(xlsChartType: int): void
        getColumn3DGroup(): excel.ChartGroup
        getCorners(): excel.Corners
        getDataTable(): excel.DataTable
        setDepthPercent(depthPercent: int): void
        getDepthPercent(): int
        setElevation(elevation: int): void
        getElevation(): int
        setGapDepth(gapDepth: int): void
        getGapDepth(): int
        setHasAxis(index1: any, index2: any, hasAxis: any): void
        isHasAxis(index1: int, index2: int): boolean
        setHasDataTable(hasDataTable: boolean): void
        isHasDataTable(): boolean
        setHasLegend(hasLegend: boolean): void
        isHasLegend(): boolean
        isHasPivotFields(): boolean
        setHeightPercent(heightPercent: int): void
        getHeightPercent(): int
        getLegend(): excel.Legend
        getLine3DGroup(): excel.ChartGroup
        getPie3DGroup(): excel.ChartGroup
        getPivotLayout(): excel.PivotLayout
        getPlotArea(): excel.PlotArea
        setPlotBy(plotBy: int): void
        getPlotBy(): int
        setProtectData(protectData: boolean): void
        isProtectData(): boolean
        isProtectionMode(): boolean
        isRightAngleAxes(): boolean
        setShowWindow(ShowWindow: boolean): void
        isShowWindow(): boolean
        isSizeWithWindow(): boolean
        getSurfaceGroup(): excel.ChartGroup
        applyCustomType(chartType: int, typeName: string): void
        applyLayout(layoutType: int, chartType: any): void
        applyDataLabels(type: int, legendKey: any, autoText: any, hasLeaderLines: any, showSeriesName: any, showCategoryName: any, showValue: any, showPercentage: any, showBubbleSize: any, separator: any): void
        areaGroups(index: int): any
        barGroups(index: int): any
        ChartObjects(index: any): any
        chartWizard(source: any, gallery: any, format: any, plotBy: any, categoryLabels: any, seriesLabels: any, hasLegend: any, title: any, categoryTitle: any, valueTitle: any, extraTitle: any): void
        lineGroups(index: int): any
        setDisplayBlanksAs(displayBlanksAs: int): void
        getDisplayBlanksAs(): int
        setHasPivotFields(hasPivotFields: boolean): void
        setPlotVisibleOnly(plotVisibleOnly: boolean): void
        isPlotVisibleOnly(): boolean
        isProtectContents(): boolean
        isProtectDrawingObjects(): boolean
        setProtectFormatting(ProtectFormatting: boolean): void
        isProtectFormatting(): boolean
        setProtectGoalSeek(ProtectGoalSeek: boolean): void
        isProtectGoalSeek(): boolean
        setProtectSelection(protectSelection: boolean): void
        isProtectSelection(): boolean
        setRightAngleAxes(rightAngleAxes: any): void
        setSizeWithWindow(SizeWithWindow: boolean): void
        setWallsAndGridlines2D(wallsAndGridlines2D: boolean): void
        isWallsAndGridlines2D(): boolean
        applyChartTemplate(templateFileName: string): void
        SaveChartTemplate(templateFileName: string): void
        getChartDataLabelInfo(): any
        setBackgroundPicture(filename: string): void
    }
    export interface Chart {
        getName(): string
        delete(): void
        setName(name: string): void
        copy(before: any, after: any): void
        location(where: int, nameObj: any): excel.Chart
        getIndex(): int
        refresh(): void
        activate(): void
        printPreview(enableChanges: any): void
        setVisible(visible: int): void
        checkSpelling(customDictionary: any, ignoreUppercase: any, alwaysSuggest: any, spellLang: any): void
        move(before: any, after: any): void
        printOut(from: any, to: any, copies: any, preview: any, activePrinter: any, printToFile: any, collate: any, prToFileName: any): void
        saveAs(filename: string, fileFormat: any, password: any, writeResPassword: any, readOnlyRecommended: any, createBackup: any, addToMru: any, textCodepage: any, textVisualLayout: any, local: any): void
        paste(type: any): void
        getPrevious(): any
        select(replace: any): void
        getPageSetup(): excel.PageSetup
        getNext(): any
        getVisible(): int
        getMailEnvelope(): int
        getScripts(): office.Scripts
        getShapes(): excel.Shapes
        unprotect(password: any): void
        exportAsFixedFormat(type: int, filename: any, quality: any, includeDocProperties: any, ignorePrintAreas: any, from: any, to: any, openAfterPublish: any, fixedFormatExtClassPtr: any): void
        protect(password: any, drawingObjects: any, contents: any, scenarios: any, userInterfaceOnly: any): void
        getCodeName(): string
        getHyperlinks(): excel.Hyperlinks
        createPublisher(edition: any, appearance: int, size: int, containsPICT: any, containsBIFF: any, containsRTF: any, containsVALU: any): void
        getRotation(): any
        setRotation(rotation: any): void
        setPerspective(perspective: int): void
        getPerspective(): int
        evaluate(name: any): any
        axes(type: any, axisGroup: int): any
        getAxes(): excel.Axes
        getFloor(): excel.Floor
        getTab(): excel.Tab
        getWalls(): excel.Walls
        deselect(): void
        export(filename: string, filterName: any, interactive: any): boolean
        xYGroups(index: int): any
        copyPicture(appearance: int, format: int, size: int): void
        columnGroups(index: int): any
        doughnutGroups(index: int): any
        getChartElement(x: int, y: int, ElementID: int, Arg1: int, Arg2: int): void
        OLEObjects(index: any): any
        pieGroups(index: int): any
        radarGroups(index: int): any
        setSourceData(source: excel.Range, plotBy: any): void
        surfaceGroups(index: int): any
        transTypeToEIO(xlsChartType: int): number[]
        transTypeToExcel(type: int, subType: int): int
        getMChart(): any
        isHasTitle(): boolean
        setHasTitle(hasTitle: boolean): void
        chartGroups(index: any): any
        seriesCollection(index: any): any
        getArea3DGroup(): excel.ChartGroup
        getChartType(): int
        setAutoScaling(autoScaling: boolean): void
        isAutoScaling(): boolean
        getBar3DGroup(): excel.ChartGroup
        getBarShape(): int
        setBarShape(barShape: int): void
        getChartArea(): excel.ChartArea
        getChartTitle(): excel.ChartTitle
        setChartStyle(chartStyle: any): void
        getChartStyle(): any
        setChartType(xlsChartType: int): void
        getColumn3DGroup(): excel.ChartGroup
        getCorners(): excel.Corners
        getDataTable(): excel.DataTable
        setDepthPercent(depthPercent: int): void
        getDepthPercent(): int
        setElevation(elevation: int): void
        getElevation(): int
        setGapDepth(gapDepth: int): void
        getGapDepth(): int
        setHasAxis(index1: any, index2: any, hasAxis: any): void
        isHasAxis(index1: int, index2: int): boolean
        setHasDataTable(hasDataTable: boolean): void
        isHasDataTable(): boolean
        setHasLegend(hasLegend: boolean): void
        isHasLegend(): boolean
        isHasPivotFields(): boolean
        setHeightPercent(heightPercent: int): void
        getHeightPercent(): int
        getLegend(): excel.Legend
        getLine3DGroup(): excel.ChartGroup
        getPie3DGroup(): excel.ChartGroup
        getPivotLayout(): excel.PivotLayout
        getPlotArea(): excel.PlotArea
        setPlotBy(plotBy: int): void
        getPlotBy(): int
        setProtectData(protectData: boolean): void
        isProtectData(): boolean
        isProtectionMode(): boolean
        isRightAngleAxes(): boolean
        setShowWindow(ShowWindow: boolean): void
        isShowWindow(): boolean
        isSizeWithWindow(): boolean
        getSurfaceGroup(): excel.ChartGroup
        applyCustomType(chartType: int, typeName: string): void
        applyLayout(layoutType: int, chartType: any): void
        applyDataLabels(type: int, legendKey: any, autoText: any, hasLeaderLines: any, showSeriesName: any, showCategoryName: any, showValue: any, showPercentage: any, showBubbleSize: any, separator: any): void
        areaGroups(index: int): any
        barGroups(index: int): any
        ChartObjects(index: any): any
        chartWizard(source: any, gallery: any, format: any, plotBy: any, categoryLabels: any, seriesLabels: any, hasLegend: any, title: any, categoryTitle: any, valueTitle: any, extraTitle: any): void
        lineGroups(index: int): any
        setDisplayBlanksAs(displayBlanksAs: int): void
        getDisplayBlanksAs(): int
        setHasPivotFields(hasPivotFields: boolean): void
        setPlotVisibleOnly(plotVisibleOnly: boolean): void
        isPlotVisibleOnly(): boolean
        isProtectContents(): boolean
        isProtectDrawingObjects(): boolean
        setProtectFormatting(ProtectFormatting: boolean): void
        isProtectFormatting(): boolean
        setProtectGoalSeek(ProtectGoalSeek: boolean): void
        isProtectGoalSeek(): boolean
        setProtectSelection(protectSelection: boolean): void
        isProtectSelection(): boolean
        setRightAngleAxes(rightAngleAxes: any): void
        setSizeWithWindow(SizeWithWindow: boolean): void
        setWallsAndGridlines2D(wallsAndGridlines2D: boolean): void
        isWallsAndGridlines2D(): boolean
        applyChartTemplate(templateFileName: string): void
        SaveChartTemplate(templateFileName: string): void
        getChartDataLabelInfo(): any
        setBackgroundPicture(filename: string): void
    }
    export interface ChartArea {
        clear(): any
        getName(): string
        copy(): any
        getBorder(): excel.Border
        getFont(): excel.Font
        getWidth(): double
        getLeft(): double
        setLeft(left: double): void
        setTop(top: double): void
        getTop(): double
        setWidth(width: double): void
        getHeight(): double
        setHeight(height: double): void
        select(): any
        isShadow(): boolean
        setShadow(shadow: boolean): void
        getFill(): excel.ChartFillFormat
        clearFormats(): any
        getInterior(): excel.Interior
        clearContents(): any
        isAutoScaleFont(): boolean
        setAutoScaleFont(autoScaleFont: boolean): void
    }
    export interface ChartArea {
        clear(): any
        getName(): string
        copy(): any
        getBorder(): excel.Border
        getFont(): excel.Font
        getWidth(): double
        getLeft(): double
        setLeft(left: double): void
        setTop(top: double): void
        getTop(): double
        setWidth(width: double): void
        getHeight(): double
        setHeight(height: double): void
        select(): any
        isShadow(): boolean
        setShadow(shadow: boolean): void
        getFill(): excel.ChartFillFormat
        clearFormats(): any
        getInterior(): excel.Interior
        clearContents(): any
        isAutoScaleFont(): boolean
        setAutoScaleFont(autoScaleFont: boolean): void
    }
    export interface ChartColorFormat {
        getType(): int
        getRGB(): int
        setRGB(rgb2: int): void
        setSchemeColor(color: int): void
        getSchemeColor(): int
    }
    export interface ChartColorFormat {
        getType(): int
        getRGB(): int
        setRGB(rgb2: int): void
        setSchemeColor(color: int): void
        getSchemeColor(): int
    }
    export interface ChartFillFormat {
        apply(mFill: any): void
        getType(): int
        setVisible(visible: int): void
        solid(): void
        getBackColor(): excel.ChartColorFormat
        getForeColor(): excel.ChartColorFormat
        getGradientStyle(): int
        getPattern(): int
        patterned(pattern: int): void
        presetGradient(style: int, variant: int, presetGradientType: int): void
        getPresetTexture(): int
        presetTextured(texture: int): void
        getTextureName(): string
        getTextureType(): int
        userPicture(pictureFile: any, pictureFormat: any, pictureStackUnit: any, picturePlacement: any): void
        userTextured(textureFile: string): void
        getVisible(): int
        oneColorGradient(style: int, variant: int, degree: double): void
        twoColorGradient(style: int, variant: int): void
        getGradientColorType(): int
        getGradientDegree(): double
        changeYzoTypeToMsType(yzoType: int): int
        changeMSTypeToYzoType(msType: int): int
        getPresetGradientType(): int
        getGradientVariant(): int
        getAngleFromMsStyle(msStyle: int): int
        setForeColor(colorFormat: excel.ChartColorFormat): void
        setBackColor(colorFormat: excel.ChartColorFormat): void
        getMFillAttribute(): any
    }
    export interface ChartFillFormat {
        apply(mFill: any): void
        getType(): int
        setVisible(visible: int): void
        solid(): void
        getBackColor(): excel.ChartColorFormat
        getForeColor(): excel.ChartColorFormat
        getGradientStyle(): int
        getPattern(): int
        patterned(pattern: int): void
        presetGradient(style: int, variant: int, presetGradientType: int): void
        getPresetTexture(): int
        presetTextured(texture: int): void
        getTextureName(): string
        getTextureType(): int
        userPicture(pictureFile: any, pictureFormat: any, pictureStackUnit: any, picturePlacement: any): void
        userTextured(textureFile: string): void
        getVisible(): int
        oneColorGradient(style: int, variant: int, degree: double): void
        twoColorGradient(style: int, variant: int): void
        getGradientColorType(): int
        getGradientDegree(): double
        changeYzoTypeToMsType(yzoType: int): int
        changeMSTypeToYzoType(msType: int): int
        getPresetGradientType(): int
        getGradientVariant(): int
        getAngleFromMsStyle(msStyle: int): int
        setForeColor(colorFormat: excel.ChartColorFormat): void
        setBackColor(colorFormat: excel.ChartColorFormat): void
        getMFillAttribute(): any
    }
    export interface ChartGroup {
        getIndex(): int
        getDropLines(): excel.DropLines
        setAxisGroup(axisGroup: int): void
        getDownBars(): excel.DownBars
        getHiLoLines(): excel.HiLoLines
        getSeriesLines(): excel.SeriesLines
        getUpBars(): excel.UpBars
        getBubbleScale(): int
        setBubbleScale(bubbleScale: int): void
        getGapWidth(): int
        setGapWidth(gapWidth: int): void
        isHas3DShading(): boolean
        setHas3DShading(has3DShading: boolean): void
        isHasDropLines(): boolean
        setHasDropLines(hasDropLines: boolean): void
        isHasHiLoLines(): boolean
        setHasHiLoLines(hasHiLoLines: boolean): void
        isHasSeriesLines(): boolean
        isHasUpDownBars(): boolean
        setHasUpDownBars(hasUpDownBars: boolean): void
        getOverlap(): int
        setOverlap(overlap: int): void
        getSplitType(): int
        setSplitType(splitType: int): void
        getSplitValue(): any
        setSplitValue(splitValue: any): void
        SeriesCollection(index: int): any
        getAxisGroup(): int
        seriesCollection(index: int): any
        getRadarAxisLabels(): excel.TickLabels
        getDoughnutHoleSize(): int
        setDoughnutHoleSize(doughnutHoleSize: int): void
        getFirstSliceAngle(): int
        setFirstSliceAngle(firstSliceAngle: int): void
        isHasRadarAxisLabels(): boolean
        setHasRadarAxisLabels(hasRadarAxisLabels: boolean): void
        setHasSeriesLines(hasSeriesLines: boolean): void
        getSecondPlotSize(): int
        setSecondPlotSize(secondPlotSize: int): void
        isShowNegativeBubbles(): boolean
        setShowNegativeBubbles(showNegativeBubbles: boolean): void
        getSizeRepresents(): int
        setSizeRepresents(sizeRepresents: int): void
        isVaryByCategories(): boolean
        setVaryByCategories(varyByCategories: boolean): void
    }
    export interface ChartGroup {
        getIndex(): int
        getDropLines(): excel.DropLines
        setAxisGroup(axisGroup: int): void
        getDownBars(): excel.DownBars
        getHiLoLines(): excel.HiLoLines
        getSeriesLines(): excel.SeriesLines
        getUpBars(): excel.UpBars
        getBubbleScale(): int
        setBubbleScale(bubbleScale: int): void
        getGapWidth(): int
        setGapWidth(gapWidth: int): void
        isHas3DShading(): boolean
        setHas3DShading(has3DShading: boolean): void
        isHasDropLines(): boolean
        setHasDropLines(hasDropLines: boolean): void
        isHasHiLoLines(): boolean
        setHasHiLoLines(hasHiLoLines: boolean): void
        isHasSeriesLines(): boolean
        isHasUpDownBars(): boolean
        setHasUpDownBars(hasUpDownBars: boolean): void
        getOverlap(): int
        setOverlap(overlap: int): void
        getSplitType(): int
        setSplitType(splitType: int): void
        getSplitValue(): any
        setSplitValue(splitValue: any): void
        SeriesCollection(index: int): any
        getAxisGroup(): int
        seriesCollection(index: int): any
        getRadarAxisLabels(): excel.TickLabels
        getDoughnutHoleSize(): int
        setDoughnutHoleSize(doughnutHoleSize: int): void
        getFirstSliceAngle(): int
        setFirstSliceAngle(firstSliceAngle: int): void
        isHasRadarAxisLabels(): boolean
        setHasRadarAxisLabels(hasRadarAxisLabels: boolean): void
        setHasSeriesLines(hasSeriesLines: boolean): void
        getSecondPlotSize(): int
        setSecondPlotSize(secondPlotSize: int): void
        isShowNegativeBubbles(): boolean
        setShowNegativeBubbles(showNegativeBubbles: boolean): void
        getSizeRepresents(): int
        setSizeRepresents(sizeRepresents: int): void
        isVaryByCategories(): boolean
        setVaryByCategories(varyByCategories: boolean): void
    }
    export interface ChartGroups {
        item(index: any): excel.ChartGroup
        getCount(): int
    }
    export interface ChartGroups {
        item(index: any): excel.ChartGroup
        getCount(): int
    }
    export interface ChartObject {
        getName(): string
        delete(): any
        setName(pName: string): void
        copy(): any
        isLocked(): boolean
        duplicate(): any
        getBorder(): excel.Border
        getIndex(): int
        getWidth(): double
        getLeft(): double
        setLeft(left: double): void
        setTop(top: double): void
        getTop(): double
        setWidth(width: double): void
        activate(): any
        getHeight(): double
        setHeight(height: double): void
        isVisible(): boolean
        setVisible(visible: boolean): void
        cut(): any
        setEnabled(enabled: boolean): void
        isEnabled(): boolean
        select(replace: any): any
        isShadow(): boolean
        setShadow(shadow: boolean): void
        getShapeRange(): excel.ShapeRange
        getMShape(): any
        setLocked(locked: boolean): void
        getHyperlinks(): excel.Hyperlinks
        getZOrder(): int
        unSelect(): void
        getInterior(): excel.Interior
        getChart(): excel.Chart
        copyPicture(appearance: int, format: int): any
        getTopLeftCell(): excel.Range
        setPlacement(placement: any): void
        getPlacement(): any
        setPrintObject(printObject: boolean): void
        isPrintObject(): boolean
        isRoundedCorners(): boolean
        bringToFront(): any
        sendToBack(): any
        getBottomRightCell(): excel.Range
        setProtectChartObject(printObject: boolean): void
        isProtectChartObject(): boolean
        setRoundedCorners(roundedCorners: boolean): void
    }
    export interface ChartObject {
        getName(): string
        delete(): any
        setName(pName: string): void
        copy(): any
        isLocked(): boolean
        duplicate(): any
        getBorder(): excel.Border
        getIndex(): int
        getWidth(): double
        getLeft(): double
        setLeft(left: double): void
        setTop(top: double): void
        getTop(): double
        setWidth(width: double): void
        activate(): any
        getHeight(): double
        setHeight(height: double): void
        isVisible(): boolean
        setVisible(visible: boolean): void
        cut(): any
        setEnabled(enabled: boolean): void
        isEnabled(): boolean
        select(replace: any): any
        isShadow(): boolean
        setShadow(shadow: boolean): void
        getShapeRange(): excel.ShapeRange
        getMShape(): any
        setLocked(locked: boolean): void
        getHyperlinks(): excel.Hyperlinks
        getZOrder(): int
        unSelect(): void
        getInterior(): excel.Interior
        getChart(): excel.Chart
        copyPicture(appearance: int, format: int): any
        getTopLeftCell(): excel.Range
        setPlacement(placement: any): void
        getPlacement(): any
        setPrintObject(printObject: boolean): void
        isPrintObject(): boolean
        isRoundedCorners(): boolean
        bringToFront(): any
        sendToBack(): any
        getBottomRightCell(): excel.Range
        setProtectChartObject(printObject: boolean): void
        isProtectChartObject(): boolean
        setRoundedCorners(roundedCorners: boolean): void
    }
    export interface ChartObjects {
        delete(): any
        copy(): any
        isLocked(): boolean
        duplicate(): any
        getBorder(): excel.Border
        item(index: any): any
        getCount(): int
        getWidth(): double
        getLeft(): double
        setLeft(left: double): void
        setTop(top: double): void
        getTop(): double
        setWidth(width: double): void
        getHeight(): double
        setHeight(height: double): void
        isVisible(): boolean
        setVisible(visible: boolean): void
        cut(): any
        setEnabled(enabled: boolean): void
        isEnabled(): boolean
        select(replace: any): any
        isShadow(): boolean
        setShadow(shadow: boolean): void
        Add(left: float, top: float, width: float, height: float): excel.ChartObject
        getShapeRange(index: any): excel.ShapeRange
        setLocked(locked: boolean): void
        getInterior(): excel.Interior
        copyPicture(appearance: int, format: int): any
        setPlacement(placement: any): void
        getPlacement(): any
        setPrintObject(printObject: boolean): void
        isPrintObject(): boolean
        isRoundedCorners(): boolean
        bringToFront(): void
        sendToBack(): void
        getChartObject(index: any): excel.ChartObject
        getChartByIndex(index: int): any
        getBottomRightCell(): excel.Range
        getAllChartObjects(): excel.ChartObject[]
        setProtectChartObject(printObject: boolean): void
        isProtectChartObject(): boolean
        setRoundedCorners(roundedCorners: boolean): void
    }
    export interface ChartObjects {
        delete(): any
        copy(): any
        isLocked(): boolean
        duplicate(): any
        getBorder(): excel.Border
        item(index: any): any
        getCount(): int
        getWidth(): double
        getLeft(): double
        setLeft(left: double): void
        setTop(top: double): void
        getTop(): double
        setWidth(width: double): void
        getHeight(): double
        setHeight(height: double): void
        isVisible(): boolean
        setVisible(visible: boolean): void
        cut(): any
        setEnabled(enabled: boolean): void
        isEnabled(): boolean
        select(replace: any): any
        isShadow(): boolean
        setShadow(shadow: boolean): void
        Add(left: float, top: float, width: float, height: float): excel.ChartObject
        getShapeRange(index: any): excel.ShapeRange
        setLocked(locked: boolean): void
        getInterior(): excel.Interior
        copyPicture(appearance: int, format: int): any
        setPlacement(placement: any): void
        getPlacement(): any
        setPrintObject(printObject: boolean): void
        isPrintObject(): boolean
        isRoundedCorners(): boolean
        bringToFront(): void
        sendToBack(): void
        getChartObject(index: any): excel.ChartObject
        getChartByIndex(index: int): any
        getBottomRightCell(): excel.Range
        getAllChartObjects(): excel.ChartObject[]
        setProtectChartObject(printObject: boolean): void
        isProtectChartObject(): boolean
        setRoundedCorners(roundedCorners: boolean): void
    }
    export interface Charts {
        add(before: excel.Worksheet, after: excel.Worksheet, count: int, type: int): excel.Chart
        delete(): void
        copy(before: any, after: any): void
        item(index: any): excel.Chart
        item(arg0: any): any
        getCount(): int
        printPreview(enableChanges: any): void
        isVisible(): boolean
        setVisible(visible: any): void
        move(before: any, after: any): void
        printOut(from: any, to: any, copies: any, preview: any, activePrinter: any, printToFile: any, collate: any, prToFileName: any): void
        printOut(from: any, to: any, copies: any, preview: any, activePrinter: any, printToFile: any, collate: any, prToFileName: any, ignorePrintAreas: any): void
        select(replace: any): void
        getHPageBreaks(): excel.HPageBreaks
        getVPageBreaks(): excel.VPageBreaks
        transTypeToEIO(xlsChartType: int): number[]
        transTypeToExcel(type: int, subType: int): int
    }
    export interface Charts {
        add(before: excel.Worksheet, after: excel.Worksheet, count: int, type: int): excel.Chart
        delete(): void
        copy(before: any, after: any): void
        item(index: any): excel.Chart
        item(arg0: any): any
        getCount(): int
        printPreview(enableChanges: any): void
        isVisible(): boolean
        setVisible(visible: any): void
        move(before: any, after: any): void
        printOut(from: any, to: any, copies: any, preview: any, activePrinter: any, printToFile: any, collate: any, prToFileName: any): void
        printOut(from: any, to: any, copies: any, preview: any, activePrinter: any, printToFile: any, collate: any, prToFileName: any, ignorePrintAreas: any): void
        select(replace: any): void
        getHPageBreaks(): excel.HPageBreaks
        getVPageBreaks(): excel.VPageBreaks
        transTypeToEIO(xlsChartType: int): number[]
        transTypeToExcel(type: int, subType: int): int
    }
    export interface ChartTitle {
        getName(): string
        delete(): any
        getBorder(): excel.Border
        getFont(): excel.Font
        getOrientation(): any
        getLeft(): double
        setLeft(left: double): void
        setTop(top: double): void
        getTop(): double
        setCaption(caption: string): void
        getCaption(): string
        getText(): string
        setText(text: string): void
        select(): any
        isShadow(): boolean
        setShadow(shadow: boolean): void
        setVerticalAlignment(verticalAlignment: any): void
        getVerticalAlignment(): any
        getFill(): excel.ChartFillFormat
        setOrientation(orientation: int): void
        getFillAttribute(): any
        getCharacters(start: any, length: any): excel.Characters
        getReadingOrder(): int
        setReadingOrder(order: any): void
        getHorizontalAlignment(): any
        setHorizontalAlignment(horizontalAlignment: any): void
        getInterior(): excel.Interior
        setFill(fillAttr: any): void
        isAutoScaleFont(): boolean
        setAutoScaleFont(autoScaleFont: boolean): void
    }
    export interface ChartTitle {
        getName(): string
        delete(): any
        getBorder(): excel.Border
        getFont(): excel.Font
        getOrientation(): any
        getLeft(): double
        setLeft(left: double): void
        setTop(top: double): void
        getTop(): double
        setCaption(caption: string): void
        getCaption(): string
        getText(): string
        setText(text: string): void
        select(): any
        isShadow(): boolean
        setShadow(shadow: boolean): void
        setVerticalAlignment(verticalAlignment: any): void
        getVerticalAlignment(): any
        getFill(): excel.ChartFillFormat
        setOrientation(orientation: int): void
        getFillAttribute(): any
        getCharacters(start: any, length: any): excel.Characters
        getReadingOrder(): int
        setReadingOrder(order: any): void
        getHorizontalAlignment(): any
        setHorizontalAlignment(horizontalAlignment: any): void
        getInterior(): excel.Interior
        setFill(fillAttr: any): void
        isAutoScaleFont(): boolean
        setAutoScaleFont(autoScaleFont: boolean): void
    }
    export interface ColorFormat {
        getType(): int
        getRGB(): int
        setRGB(rgb2: int): void
        getTintAndShade(): double
        setTintAndShade(tintAndShade: double): void
        getBrightness(): double
        setBrightness(bright: double): void
        setSchemeColor(color: int): void
        getSchemeColor(): int
    }
    export interface ColorFormat {
        getType(): int
        getRGB(): int
        setRGB(rgb2: int): void
        getTintAndShade(): double
        setTintAndShade(tintAndShade: double): void
        getBrightness(): double
        setBrightness(bright: double): void
        setSchemeColor(color: int): void
        getSchemeColor(): int
    }
    export interface ColorScale {
        delete(): void
        setPriority(p: int): void
        getPriority(): int
        getType(): int
        setFormula(f: string): void
        appliesTo(): excel.Range
        setFirstPriority(): void
        setLastPriority(): void
        getStopIfTrue(): boolean
        setStopIfTrue(s: boolean): void
        modifyAppliesToRange(range: excel.Range): void
        getPTCondition(): boolean
        getScopeType(): int
        setScopeType(pcs: int): void
        getColorScaleCriteria(): excel.ColorScaleCriteria
        getFormula(): string
    }
    export interface ColorScale {
        delete(): void
        setPriority(p: int): void
        getPriority(): int
        getType(): int
        setFormula(f: string): void
        appliesTo(): excel.Range
        setFirstPriority(): void
        setLastPriority(): void
        getStopIfTrue(): boolean
        setStopIfTrue(s: boolean): void
        modifyAppliesToRange(range: excel.Range): void
        getPTCondition(): boolean
        getScopeType(): int
        setScopeType(pcs: int): void
        getColorScaleCriteria(): excel.ColorScaleCriteria
        getFormula(): string
    }
    export interface ColorScaleCriteria {
        item(i: any): excel.ColorScaleCriterion
        getCount(): int
    }
    export interface ColorScaleCriteria {
        item(i: any): excel.ColorScaleCriterion
        getCount(): int
    }
    export interface ColorScaleCriterion {
        getValue(): any
        setValue(value: any): void
        getType(): int
        getIndex(): int
        setType(type: int): void
        getFormatColor(): excel.FormatColor
        getMColorScaleCriterion(): any
    }
    export interface ColorScaleCriterion {
        getValue(): any
        setValue(value: any): void
        getType(): int
        getIndex(): int
        setType(type: int): void
        getFormatColor(): excel.FormatColor
        getMColorScaleCriterion(): any
    }
    export interface Comment {
        next(): excel.Comment
        delete(): void
        previous(): excel.Comment
        text(text: any, start: any, overwrite: any): string
        getShape(): excel.Shape
        isVisible(): boolean
        setVisible(visible: boolean): void
        getText(): string
        getAuthor(): string
        setText2(text: string, start: int, overwrite: boolean): void
        pasteText(): string[]
    }
    export interface Comment {
        next(): excel.Comment
        delete(): void
        previous(): excel.Comment
        text(text: any, start: any, overwrite: any): string
        getShape(): excel.Shape
        isVisible(): boolean
        setVisible(visible: boolean): void
        getText(): string
        getAuthor(): string
        setText2(text: string, start: int, overwrite: boolean): void
        pasteText(): string[]
    }
    export interface Comments {
        item(index: int): excel.Comment
        getCount(): int
    }
    export interface Comments {
        item(index: int): excel.Comment
        getCount(): int
    }
    export interface ConditionValue {
        getValue(): any
        setValue(value: any): void
        getType(): int
        modify(type: int, v: any): void
    }
    export interface ConditionValue {
        getValue(): any
        setValue(value: any): void
        getType(): int
        modify(type: int, v: any): void
    }
    export interface ConnectorFormat {
        getType(): int
        setType(type: int): void
        isBeginConnected(): boolean
        isEndConnected(): boolean
        beginConnect(connectedShape: excel.Shape, connectionSite: int): void
        beginDisconnect(): void
        endConnect(connectedShape: excel.Shape, connectionSite: int): void
        endDisconnect(): void
        getBeginConnectedShape(): excel.Shape
        getBeginConnectionSite(): int
        getEndConnectedShape(): excel.Shape
        getEndConnectionSite(): int
    }
    export interface ConnectorFormat {
        getType(): int
        setType(type: int): void
        isBeginConnected(): boolean
        isEndConnected(): boolean
        beginConnect(connectedShape: excel.Shape, connectionSite: int): void
        beginDisconnect(): void
        endConnect(connectedShape: excel.Shape, connectionSite: int): void
        endDisconnect(): void
        getBeginConnectedShape(): excel.Shape
        getBeginConnectionSite(): int
        getEndConnectedShape(): excel.Shape
        getEndConnectionSite(): int
    }
    export interface ControlFormat {
        getValue(): int
        setValue(value: int): void
        setEnabled(enabled: boolean): void
        isEnabled(): boolean
        getList(index: any): any
        setPrintObject(printObject: boolean): void
        isPrintObject(): boolean
        getDropDownLines(): int
        setDropDownLines(dropDownLines: int): void
        getLargeChange(): int
        setLargeChange(largeChange: int): void
        getLinkedCell(): string
        setLinkedCell(linkedCell: string): void
        getListCount(): int
        setListCount(listCount: int): void
        getListFillRange(): string
        setListFillRange(listFillRange: string): void
        getListIndex(): int
        setListIndex(listIndex: int): void
        isLockedText(): boolean
        setLockedText(lockedText: boolean): void
        getMultiSelect(): int
        setMultiSelect(multiSelect: int): void
        getSmallChange(): int
        setSmallChange(smallChange: int): void
        removeAllItems(): void
        removeItem(index: int, count: any): void
        getMax(): int
        setMax(max: int): void
        getMin(): int
        setMin(min: int): void
        addItem(text: string, index: any): void
        setList(list: any): void
    }
    export interface ControlFormat {
        getValue(): int
        setValue(value: int): void
        setEnabled(enabled: boolean): void
        isEnabled(): boolean
        getList(index: any): any
        setPrintObject(printObject: boolean): void
        isPrintObject(): boolean
        getDropDownLines(): int
        setDropDownLines(dropDownLines: int): void
        getLargeChange(): int
        setLargeChange(largeChange: int): void
        getLinkedCell(): string
        setLinkedCell(linkedCell: string): void
        getListCount(): int
        setListCount(listCount: int): void
        getListFillRange(): string
        setListFillRange(listFillRange: string): void
        getListIndex(): int
        setListIndex(listIndex: int): void
        isLockedText(): boolean
        setLockedText(lockedText: boolean): void
        getMultiSelect(): int
        setMultiSelect(multiSelect: int): void
        getSmallChange(): int
        setSmallChange(smallChange: int): void
        removeAllItems(): void
        removeItem(index: int, count: any): void
        getMax(): int
        setMax(max: int): void
        getMin(): int
        setMin(min: int): void
        addItem(text: string, index: any): void
        setList(list: any): void
    }
    export interface Corners {
        getName(): string
        select(): void
    }
    export interface Corners {
        getName(): string
        select(): void
    }
    export interface CubeField {
        getName(): string
        getValue(): string
        delete(): void
        getOrientation(): int
        getCaption(): string
        setPosition(position: int): void
        getPosition(): int
        setOrientation(orientation: int): void
        isEnableMultiplePageItems(): boolean
        setEnableMultiplePageItems(enableMultiplePageItems: boolean): void
        getLayoutSubtotalLocation(): int
        setLayoutSubtotalLocation(layoutSubtotalLocation: int): void
        getCubeFieldType(): int
        setDragToColumn(dragToColumn: boolean): void
        isDragToColumn(): boolean
        isDragToData(): boolean
        setDragToData(dragToData: boolean): void
        isDragToHide(): boolean
        setDragToHide(dragToHide: boolean): void
        isDragToPage(): boolean
        setDragToPage(dragToPage: boolean): void
        isDragToRow(): boolean
        setDragToRow(dragToRow: boolean): void
        getHiddenLevels(): int
        setHiddenLevels(hiddenLevels: int): void
        getLayoutForm(): int
        setLayoutForm(layoutForm: int): void
        getPivotFields(): excel.PivotFields
        isHasMemberProperties(): boolean
        isShowInFieldList(): boolean
        setShowInFieldList(showInFieldList: boolean): void
        getTreeviewControl(): excel.TreeviewControl
        addMemberPropertyField(property: string, propertyOrder: int): void
    }
    export interface CubeField {
        getName(): string
        getValue(): string
        delete(): void
        getOrientation(): int
        getCaption(): string
        setPosition(position: int): void
        getPosition(): int
        setOrientation(orientation: int): void
        isEnableMultiplePageItems(): boolean
        setEnableMultiplePageItems(enableMultiplePageItems: boolean): void
        getLayoutSubtotalLocation(): int
        setLayoutSubtotalLocation(layoutSubtotalLocation: int): void
        getCubeFieldType(): int
        setDragToColumn(dragToColumn: boolean): void
        isDragToColumn(): boolean
        isDragToData(): boolean
        setDragToData(dragToData: boolean): void
        isDragToHide(): boolean
        setDragToHide(dragToHide: boolean): void
        isDragToPage(): boolean
        setDragToPage(dragToPage: boolean): void
        isDragToRow(): boolean
        setDragToRow(dragToRow: boolean): void
        getHiddenLevels(): int
        setHiddenLevels(hiddenLevels: int): void
        getLayoutForm(): int
        setLayoutForm(layoutForm: int): void
        getPivotFields(): excel.PivotFields
        isHasMemberProperties(): boolean
        isShowInFieldList(): boolean
        setShowInFieldList(showInFieldList: boolean): void
        getTreeviewControl(): excel.TreeviewControl
        addMemberPropertyField(property: string, propertyOrder: int): void
    }
    export interface CubeFields {
        item(index: any): excel.CubeField
        getCount(): int
        addSet(name: string, caption: string): excel.CubeField
    }
    export interface CubeFields {
        item(index: any): excel.CubeField
        getCount(): int
        addSet(name: string, caption: string): excel.CubeField
    }
    export interface CustomProperties {
        add(name: string, value: any): excel.CustomProperty
        item(index: any): excel.CustomProperty
        getCount(): int
    }
    export interface CustomProperties {
        add(name: string, value: any): excel.CustomProperty
        item(index: any): excel.CustomProperty
        getCount(): int
    }
    export interface CustomProperty {
        getName(): string
        getValue(): any
        delete(): void
        setName(name: string): void
        setValue(value: any): void
    }
    export interface CustomProperty {
        getName(): string
        getValue(): any
        delete(): void
        setName(name: string): void
        setValue(value: any): void
    }
    export interface CustomView {
        getName(): string
        delete(): void
        show(): void
        isPrintSettings(): boolean
        isRowColSettings(): boolean
    }
    export interface CustomView {
        getName(): string
        delete(): void
        show(): void
        isPrintSettings(): boolean
        isRowColSettings(): boolean
    }
    export interface CustomViews {
        add(viewName: string, printSettings: boolean, rowColSettings: boolean): excel.CustomView
        item(view: any): excel.CustomView
        getCount(): int
    }
    export interface CustomViews {
        add(viewName: string, printSettings: boolean, rowColSettings: boolean): excel.CustomView
        item(view: any): excel.CustomView
        getCount(): int
    }
    export interface Databar {
        delete(): void
        setPriority(p: int): void
        getPriority(): int
        getType(): int
        setFormula(ap: string): void
        getDirection(): int
        setDirection(ap: int): void
        getAppliesTo(): excel.Range
        setFirstPriority(): void
        setLastPriority(): void
        getStopIfTrue(): boolean
        setStopIfTrue(s: boolean): void
        modifyAppliesToRange(range: excel.Range): void
        getMDatabar(): any
        getAxisColor(): any
        getAxisPosition(): int
        setAxisPosition(ap: int): void
        getDataBarBorder(): excel.DataBarBorder
        getBarColor(): any
        getBarFillType(): int
        setBarFillType(ap: int): void
        getMaxPoint(): excel.ConditionValue
        getMinPoint(): excel.ConditionValue
        getPercentMax(): int
        setPercentMax(pm: int): void
        getPercentMin(): int
        setPercentMin(pm: int): void
        getShowValue(): boolean
        getFormula(): string
        setShowValue(sv: boolean): void
        getNegativeBarFormat(): excel.NegativeBarFormat
    }
    export interface Databar {
        delete(): void
        setPriority(p: int): void
        getPriority(): int
        getType(): int
        setFormula(ap: string): void
        getDirection(): int
        setDirection(ap: int): void
        getAppliesTo(): excel.Range
        setFirstPriority(): void
        setLastPriority(): void
        getStopIfTrue(): boolean
        setStopIfTrue(s: boolean): void
        modifyAppliesToRange(range: excel.Range): void
        getMDatabar(): any
        getAxisColor(): any
        getAxisPosition(): int
        setAxisPosition(ap: int): void
        getDataBarBorder(): excel.DataBarBorder
        getBarColor(): any
        getBarFillType(): int
        setBarFillType(ap: int): void
        getMaxPoint(): excel.ConditionValue
        getMinPoint(): excel.ConditionValue
        getPercentMax(): int
        setPercentMax(pm: int): void
        getPercentMin(): int
        setPercentMin(pm: int): void
        getShowValue(): boolean
        getFormula(): string
        setShowValue(sv: boolean): void
        getNegativeBarFormat(): excel.NegativeBarFormat
    }
    export interface DataBarBorder {
        getType(): int
        setType(type: int): void
        getBarColor(): excel.FormatColor
    }
    export interface DataBarBorder {
        getType(): int
        setType(type: int): void
        getBarColor(): excel.FormatColor
    }
    export interface DataLabel {
        getName(): string
        delete(): any
        getType(): int
        getSeparator(): any
        getBorder(): excel.Border
        getFont(): excel.Font
        setType(type: int): void
        getOrientation(): any
        getLeft(): double
        setLeft(f: double): void
        setTop(f: double): void
        getTop(): double
        setCaption(fontName: string): void
        getCaption(): string
        getText(): string
        setText(fontName: string): void
        select(): any
        isShadow(): boolean
        setPosition(l: int): void
        getPosition(): int
        setSeparator(fontName: any): void
        setShadow(flag: boolean): void
        setVerticalAlignment(l: any): void
        getVerticalAlignment(): any
        getFill(): excel.ChartFillFormat
        setOrientation(s: any): void
        getNumberFormat(): string
        setNumberFormat(s: string): void
        getCharacters(start: any, length: any): excel.Characters
        getReadingOrder(): int
        setReadingOrder(l: int): void
        getHorizontalAlignment(): any
        setHorizontalAlignment(l: any): void
        setAutoText(flag: boolean): void
        getInterior(): excel.Interior
        isShowValue(): boolean
        isAutoText(): boolean
        isShowBubbleSize(): boolean
        setShowLegendKey(flag: boolean): void
        isShowLegendKey(): boolean
        isShowPercentage(): boolean
        isShowSeriesName(): boolean
        isAutoScaleFont(): boolean
        setAutoScaleFont(flag: boolean): void
        setShowValue(flag: boolean): void
        setNumberFormatLocal(s: any): void
        getNumberFormatLocal(): any
        setShowSeriesName(flag: boolean): void
        setShowCategoryName(flag: boolean): void
        setShowPercentage(flag: boolean): void
        setShowBubbleSize(flag: boolean): void
        setNumberFormatLinked(flag: boolean): void
        isNumberFormatLinked(): boolean
        isShowCategoryName(): boolean
    }
    export interface DataLabel {
        getName(): string
        delete(): any
        getType(): int
        getSeparator(): any
        getBorder(): excel.Border
        getFont(): excel.Font
        setType(type: int): void
        getOrientation(): any
        getLeft(): double
        setLeft(f: double): void
        setTop(f: double): void
        getTop(): double
        setCaption(fontName: string): void
        getCaption(): string
        getText(): string
        setText(fontName: string): void
        select(): any
        isShadow(): boolean
        setPosition(l: int): void
        getPosition(): int
        setSeparator(fontName: any): void
        setShadow(flag: boolean): void
        setVerticalAlignment(l: any): void
        getVerticalAlignment(): any
        getFill(): excel.ChartFillFormat
        setOrientation(s: any): void
        getNumberFormat(): string
        setNumberFormat(s: string): void
        getCharacters(start: any, length: any): excel.Characters
        getReadingOrder(): int
        setReadingOrder(l: int): void
        getHorizontalAlignment(): any
        setHorizontalAlignment(l: any): void
        setAutoText(flag: boolean): void
        getInterior(): excel.Interior
        isShowValue(): boolean
        isAutoText(): boolean
        isShowBubbleSize(): boolean
        setShowLegendKey(flag: boolean): void
        isShowLegendKey(): boolean
        isShowPercentage(): boolean
        isShowSeriesName(): boolean
        isAutoScaleFont(): boolean
        setAutoScaleFont(flag: boolean): void
        setShowValue(flag: boolean): void
        setNumberFormatLocal(s: any): void
        getNumberFormatLocal(): any
        setShowSeriesName(flag: boolean): void
        setShowCategoryName(flag: boolean): void
        setShowPercentage(flag: boolean): void
        setShowBubbleSize(flag: boolean): void
        setNumberFormatLinked(flag: boolean): void
        isNumberFormatLinked(): boolean
        isShowCategoryName(): boolean
    }
    export interface DataLabels {
        getName(): string
        delete(): any
        getType(): int
        getSeparator(): any
        getBorder(): excel.Border
        item(index: any): excel.DataLabel
        getFont(): excel.Font
        setType(type: int): void
        getOrientation(): any
        getCount(): int
        select(): any
        isShadow(): boolean
        setPosition(l: int): void
        getPosition(): int
        setSeparator(fontName: any): void
        setShadow(flag: boolean): void
        setVerticalAlignment(l: any): void
        getVerticalAlignment(): any
        getFill(): excel.FillFormat
        setOrientation(s: any): void
        getNumberFormat(): string
        setNumberFormat(s: string): void
        getReadingOrder(): int
        setReadingOrder(l: int): void
        getHorizontalAlignment(): any
        setHorizontalAlignment(l: any): void
        setAutoText(flag: boolean): void
        getInterior(): excel.Interior
        isShowValue(): boolean
        isAutoText(): boolean
        isShowBubbleSize(): boolean
        setShowLegendKey(flag: boolean): void
        isShowLegendKey(): boolean
        isShowPercentage(): boolean
        isShowSeriesName(): boolean
        isAutoScaleFont(): boolean
        setAutoScaleFont(flag: boolean): void
        setShowValue(flag: boolean): void
        setNumberFormatLocal(s: any): void
        getNumberFormatLocal(): any
        setShowSeriesName(flag: boolean): void
        setShowCategoryName(flag: boolean): void
        setShowPercentage(flag: boolean): void
        setShowBubbleSize(flag: boolean): void
        setNumberFormatLinked(flag: boolean): void
        isNumberFormatLinked(): boolean
        isShowCategoryName(): boolean
    }
    export interface DataLabels {
        getName(): string
        delete(): any
        getType(): int
        getSeparator(): any
        getBorder(): excel.Border
        item(index: any): excel.DataLabel
        getFont(): excel.Font
        setType(type: int): void
        getOrientation(): any
        getCount(): int
        select(): any
        isShadow(): boolean
        setPosition(l: int): void
        getPosition(): int
        setSeparator(fontName: any): void
        setShadow(flag: boolean): void
        setVerticalAlignment(l: any): void
        getVerticalAlignment(): any
        getFill(): excel.FillFormat
        setOrientation(s: any): void
        getNumberFormat(): string
        setNumberFormat(s: string): void
        getReadingOrder(): int
        setReadingOrder(l: int): void
        getHorizontalAlignment(): any
        setHorizontalAlignment(l: any): void
        setAutoText(flag: boolean): void
        getInterior(): excel.Interior
        isShowValue(): boolean
        isAutoText(): boolean
        isShowBubbleSize(): boolean
        setShowLegendKey(flag: boolean): void
        isShowLegendKey(): boolean
        isShowPercentage(): boolean
        isShowSeriesName(): boolean
        isAutoScaleFont(): boolean
        setAutoScaleFont(flag: boolean): void
        setShowValue(flag: boolean): void
        setNumberFormatLocal(s: any): void
        getNumberFormatLocal(): any
        setShowSeriesName(flag: boolean): void
        setShowCategoryName(flag: boolean): void
        setShowPercentage(flag: boolean): void
        setShowBubbleSize(flag: boolean): void
        setNumberFormatLinked(flag: boolean): void
        isNumberFormatLinked(): boolean
        isShowCategoryName(): boolean
    }
    export interface DataTable {
        delete(): void
        getBorder(): excel.Border
        getFont(): excel.Font
        select(): void
        setShowLegendKey(flag: boolean): void
        isShowLegendKey(): boolean
        hasBorderOutline(): boolean
        isAutoScaleFont(): boolean
        setAutoScaleFont(flag: boolean): void
        setHasBorderHorizontal(flag: boolean): void
        hasBorderHorizontal(): boolean
        setHasBorderVertical(flag: boolean): void
        hasBorderVertical(): boolean
        setHasBorderOutline(flag: boolean): void
    }
    export interface DataTable {
        delete(): void
        getBorder(): excel.Border
        getFont(): excel.Font
        select(): void
        setShowLegendKey(flag: boolean): void
        isShowLegendKey(): boolean
        hasBorderOutline(): boolean
        isAutoScaleFont(): boolean
        setAutoScaleFont(flag: boolean): void
        setHasBorderHorizontal(flag: boolean): void
        hasBorderHorizontal(): boolean
        setHasBorderVertical(flag: boolean): void
        hasBorderVertical(): boolean
        setHasBorderOutline(flag: boolean): void
    }
    export interface DefaultWebOptions {
        getEncoding(): int
        isOrganizeInFolder(): boolean
        setOrganizeInFolder(flag: boolean): void
        setUpdateLinksOnSave(flag: boolean): void
        isUpdateLinksOnSave(): boolean
        setUseLongFileNames(flag: boolean): void
        isUseLongFileNames(): boolean
        getFolderSuffix(): string
        isAllowPNG(): boolean
        setAllowPNG(flag: boolean): void
        setEncoding(l: int): void
        isAlwaysSaveInDefaultEncoding(): boolean
        setAlwaysSaveInDefaultEncoding(flag: boolean): void
        isCheckIfOfficeIsHTMLEditor(): boolean
        setCheckIfOfficeIsHTMLEditor(flag: boolean): void
        isSaveNewWebPagesAsWebArchives(): boolean
        setSaveNewWebPagesAsWebArchives(flag: boolean): void
        getPixelsPerInch(): int
        setPixelsPerInch(l: int): void
        isRelyOnCSS(): boolean
        setRelyOnCSS(flag: boolean): void
        isRelyOnVML(): boolean
        setRelyOnVML(flag: boolean): void
        getScreenSize(): int
        setScreenSize(l: int): void
        setTargetBrowser(l: int): void
        getTargetBrowser(): int
        getFonts(): office.WebPageFonts
        setLoadPictures(flag: boolean): void
        isLoadPictures(): boolean
        isSaveHiddenData(): boolean
        setDownloadComponents(flag: boolean): void
        isDownloadComponents(): boolean
        getLocationOfComponents(): string
        setLocationOfComponents(s: string): void
        setSaveHiddenData(flag: boolean): void
    }
    export interface DefaultWebOptions {
        getEncoding(): int
        isOrganizeInFolder(): boolean
        setOrganizeInFolder(flag: boolean): void
        setUpdateLinksOnSave(flag: boolean): void
        isUpdateLinksOnSave(): boolean
        setUseLongFileNames(flag: boolean): void
        isUseLongFileNames(): boolean
        getFolderSuffix(): string
        isAllowPNG(): boolean
        setAllowPNG(flag: boolean): void
        setEncoding(l: int): void
        isAlwaysSaveInDefaultEncoding(): boolean
        setAlwaysSaveInDefaultEncoding(flag: boolean): void
        isCheckIfOfficeIsHTMLEditor(): boolean
        setCheckIfOfficeIsHTMLEditor(flag: boolean): void
        isSaveNewWebPagesAsWebArchives(): boolean
        setSaveNewWebPagesAsWebArchives(flag: boolean): void
        getPixelsPerInch(): int
        setPixelsPerInch(l: int): void
        isRelyOnCSS(): boolean
        setRelyOnCSS(flag: boolean): void
        isRelyOnVML(): boolean
        setRelyOnVML(flag: boolean): void
        getScreenSize(): int
        setScreenSize(l: int): void
        setTargetBrowser(l: int): void
        getTargetBrowser(): int
        getFonts(): office.WebPageFonts
        setLoadPictures(flag: boolean): void
        isLoadPictures(): boolean
        isSaveHiddenData(): boolean
        setDownloadComponents(flag: boolean): void
        isDownloadComponents(): boolean
        getLocationOfComponents(): string
        setLocationOfComponents(s: string): void
        setSaveHiddenData(flag: boolean): void
    }
    export interface Diagram {
        getType(): int
        convert(type: int): void
        setAutoFormat(l: int): void
        setAutoLayout(l: int): void
        setReverse(l: int): void
        fitText(): void
        getNodes(): excel.DiagramNodes
        getAutoFormat(): int
        getAutoLayout(): int
        getReverse(): int
    }
    export interface Diagram {
        getType(): int
        convert(type: int): void
        setAutoFormat(l: int): void
        setAutoLayout(l: int): void
        setReverse(l: int): void
        fitText(): void
        getNodes(): excel.DiagramNodes
        getAutoFormat(): int
        getAutoLayout(): int
        getReverse(): int
    }
    export interface DiagramNode {
        delete(): void
        getRoot(): excel.DiagramNode
        replaceNode(pTargetNode: excel.DiagramNode): void
        setLayout(l: int): void
        getLayout(): int
        getShape(): excel.Shape
        nextNode(): excel.DiagramNode
        prevNode(): excel.DiagramNode
        getChildren(): excel.DiagramNodeChildren
        getDiagram(): excel.Diagram
        getTextShape(): excel.Shape
        addNode(pos: int, nodeType: int): excel.DiagramNode
        CloneNode(copyChildren: boolean, pTargetNode: excel.DiagramNode, pos: int): excel.DiagramNode
        TransferChildren(pReceivingNode: excel.DiagramNode): void
        MoveNode(pTargetNode: excel.DiagramNode, pos: int): void
        SwapNode(pTargetNode: excel.DiagramNode, swapChildren: boolean): void
    }
    export interface DiagramNode {
        delete(): void
        getRoot(): excel.DiagramNode
        replaceNode(pTargetNode: excel.DiagramNode): void
        setLayout(l: int): void
        getLayout(): int
        getShape(): excel.Shape
        nextNode(): excel.DiagramNode
        prevNode(): excel.DiagramNode
        getChildren(): excel.DiagramNodeChildren
        getDiagram(): excel.Diagram
        getTextShape(): excel.Shape
        addNode(pos: int, nodeType: int): excel.DiagramNode
        CloneNode(copyChildren: boolean, pTargetNode: excel.DiagramNode, pos: int): excel.DiagramNode
        TransferChildren(pReceivingNode: excel.DiagramNode): void
        MoveNode(pTargetNode: excel.DiagramNode, pos: int): void
        SwapNode(pTargetNode: excel.DiagramNode, swapChildren: boolean): void
    }
    export interface DiagramNodeChildren {
        getCount(): int
        Item(index: int): excel.DiagramNode
        selectAll(): void
        getFirstChild(): excel.DiagramNode
        getLastChild(): excel.DiagramNode
        addNode(Index: int, index: int): excel.DiagramNode
    }
    export interface DiagramNodeChildren {
        getCount(): int
        Item(index: int): excel.DiagramNode
        selectAll(): void
        getFirstChild(): excel.DiagramNode
        getLastChild(): excel.DiagramNode
        addNode(Index: int, index: int): excel.DiagramNode
    }
    export interface DiagramNodes {
        item(index: int): excel.DiagramNode
        getCount(): int
        selectAll(): void
    }
    export interface DiagramNodes {
        item(index: int): excel.DiagramNode
        getCount(): int
        selectAll(): void
    }
    export interface Dialog {
        show(args1: any, args2: any, args3: any, args4: any, args5: any, args6: any, args7: any, args8: any, args9: any, args10: any, args11: any, args12: any, args13: any, args14: any, args15: any, args16: any, args17: any, args18: any, args19: any, args20: any, args21: any, args22: any, args23: any, args24: any, args25: any, args26: any, args27: any, args28: any, args29: any, args30: any): boolean
        show(args1: any): boolean
        show2(args1: any, args2: any, args3: any, args4: any, args5: any, args6: any, args7: any, args8: any, args9: any, args10: any, args11: any, args12: any, args13: any, args14: any, args15: any, args16: any, args17: any, args18: any, args19: any, args20: any, args21: any, args22: any, args23: any, args24: any, args25: any, args26: any, args27: any, args28: any, args29: any, args30: any, callBack: boolean): boolean
        show2(args1: any, args2: any, args3: any, args4: any, args5: any, args6: any, args7: any, args8: any, args9: any, args10: any, args11: any, args12: any, args13: any, args14: any, args15: any, args16: any, args17: any, args18: any, args19: any, args20: any, args21: any, args22: any, args23: any, args24: any, args25: any, args26: any, args27: any, args28: any, args29: any, args30: any): boolean
        show2(args1: any): boolean
    }
    export interface Dialog {
        show(args1: any, args2: any, args3: any, args4: any, args5: any, args6: any, args7: any, args8: any, args9: any, args10: any, args11: any, args12: any, args13: any, args14: any, args15: any, args16: any, args17: any, args18: any, args19: any, args20: any, args21: any, args22: any, args23: any, args24: any, args25: any, args26: any, args27: any, args28: any, args29: any, args30: any): boolean
        show(args1: any): boolean
        show2(args1: any, args2: any, args3: any, args4: any, args5: any, args6: any, args7: any, args8: any, args9: any, args10: any, args11: any, args12: any, args13: any, args14: any, args15: any, args16: any, args17: any, args18: any, args19: any, args20: any, args21: any, args22: any, args23: any, args24: any, args25: any, args26: any, args27: any, args28: any, args29: any, args30: any, callBack: boolean): boolean
        show2(args1: any, args2: any, args3: any, args4: any, args5: any, args6: any, args7: any, args8: any, args9: any, args10: any, args11: any, args12: any, args13: any, args14: any, args15: any, args16: any, args17: any, args18: any, args19: any, args20: any, args21: any, args22: any, args23: any, args24: any, args25: any, args26: any, args27: any, args28: any, args29: any, args30: any): boolean
        show2(args1: any): boolean
    }
    export interface Dialogs {
        item(index: int): excel.Dialog
        getCount(): int
    }
    export interface Dialogs {
        item(index: int): excel.Dialog
        getCount(): int
    }
    export interface DisplayFormat {
        getFont(): excel.Font
        getOrientation(): any
        getBorders(): excel.Borders
        getVerticalAlignment(): any
        getLocked(): any
        getNumberFormat(): any
        getStyle(): any
        getCharacters(start: any, length: any): excel.Characters
        getReadingOrder(): int
        getHorizontalAlignment(): any
        getInterior(): excel.Interior
        getAddIndent(): any
        getFormulaHidden(): any
        getShrinkToFit(): any
        getWrapText(): any
        getIndentLevel(): any
        isMergeCells(): any
        getNumberFormatLocal(): any
    }
    export interface DisplayFormat {
        getFont(): excel.Font
        getOrientation(): any
        getBorders(): excel.Borders
        getVerticalAlignment(): any
        getLocked(): any
        getNumberFormat(): any
        getStyle(): any
        getCharacters(start: any, length: any): excel.Characters
        getReadingOrder(): int
        getHorizontalAlignment(): any
        getInterior(): excel.Interior
        getAddIndent(): any
        getFormulaHidden(): any
        getShrinkToFit(): any
        getWrapText(): any
        getIndentLevel(): any
        isMergeCells(): any
        getNumberFormatLocal(): any
    }
    export interface DisplayUnitLabel {
        getName(): string
        delete(): any
        getBorder(): excel.Border
        getFont(): excel.Font
        getOrientation(): any
        getLeft(): double
        setLeft(f: double): void
        setTop(f: double): void
        getTop(): double
        setCaption(fontName: string): void
        getCaption(): string
        getText(): string
        setText(fontName: string): void
        select(): any
        isShadow(): boolean
        setShadow(flag: boolean): void
        setVerticalAlignment(l: any): void
        getVerticalAlignment(): any
        getFill(): excel.FillFormat
        setOrientation(s: any): void
        getCharacters(start: any, length: any): excel.Characters
        getReadingOrder(): int
        setReadingOrder(l: int): void
        getHorizontalAlignment(): any
        setHorizontalAlignment(l: any): void
        getInterior(): excel.Interior
        isAutoScaleFont(): boolean
        setAutoScaleFont(flag: boolean): void
    }
    export interface DisplayUnitLabel {
        getName(): string
        delete(): any
        getBorder(): excel.Border
        getFont(): excel.Font
        getOrientation(): any
        getLeft(): double
        setLeft(f: double): void
        setTop(f: double): void
        getTop(): double
        setCaption(fontName: string): void
        getCaption(): string
        getText(): string
        setText(fontName: string): void
        select(): any
        isShadow(): boolean
        setShadow(flag: boolean): void
        setVerticalAlignment(l: any): void
        getVerticalAlignment(): any
        getFill(): excel.FillFormat
        setOrientation(s: any): void
        getCharacters(start: any, length: any): excel.Characters
        getReadingOrder(): int
        setReadingOrder(l: int): void
        getHorizontalAlignment(): any
        setHorizontalAlignment(l: any): void
        getInterior(): excel.Interior
        isAutoScaleFont(): boolean
        setAutoScaleFont(flag: boolean): void
    }
    export interface DownBars {
        getName(): string
        delete(): any
        getBorder(): excel.Border
        select(): any
        getFill(): excel.FillFormat
        getInterior(): excel.Interior
    }
    export interface DownBars {
        getName(): string
        delete(): any
        getBorder(): excel.Border
        select(): any
        getFill(): excel.FillFormat
        getInterior(): excel.Interior
    }
    export interface DrawingObjects {
        delete(): void
        getShapeRange(): excel.ShapeRange
    }
    export interface DrawingObjects {
        delete(): void
        getShapeRange(): excel.ShapeRange
    }
    export interface DropLines {
        getName(): string
        delete(): any
        getBorder(): excel.Border
        select(): any
    }
    export interface DropLines {
        getName(): string
        delete(): any
        getBorder(): excel.Border
        select(): any
    }
    export interface Error {
        getValue(): boolean
        setIgnore(b: boolean): void
        isIgnore(): boolean
    }
    export interface Error {
        getValue(): boolean
        setIgnore(b: boolean): void
        isIgnore(): boolean
    }
    export interface ErrorBars {
        getName(): string
        delete(): any
        getBorder(): excel.Border
        select(): any
        clearFormats(): any
        getEndStyle(): int
        setEndStyle(l: int): void
    }
    export interface ErrorBars {
        getName(): string
        delete(): any
        getBorder(): excel.Border
        select(): any
        clearFormats(): any
        getEndStyle(): int
        setEndStyle(l: int): void
    }
    export interface ErrorCheckingOptions {
        setNumberAsText(flag: boolean): void
        isNumberAsText(): boolean
        setOmittedCells(flag: boolean): void
        isOmittedCells(): boolean
        setTextDate(flag: boolean): void
        isTextDate(): boolean
        setBackgroundChecking(flag: boolean): void
        isBackgroundChecking(): boolean
        setEmptyCellReferences(flag: boolean): void
        isEmptyCellReferences(): boolean
        setEvaluateToError(flag: boolean): void
        isEvaluateToError(): boolean
        setInconsistentFormula(flag: boolean): void
        isInconsistentFormula(): boolean
        setIndicatorColorIndex(l: int): void
        IndicatorColorIndex(): int
        setListDataValidation(flag: boolean): void
        isListDataValidation(): boolean
        setUnlockedFormulaCells(flag: boolean): void
        isUnlockedFormulaCells(): boolean
    }
    export interface ErrorCheckingOptions {
        setNumberAsText(flag: boolean): void
        isNumberAsText(): boolean
        setOmittedCells(flag: boolean): void
        isOmittedCells(): boolean
        setTextDate(flag: boolean): void
        isTextDate(): boolean
        setBackgroundChecking(flag: boolean): void
        isBackgroundChecking(): boolean
        setEmptyCellReferences(flag: boolean): void
        isEmptyCellReferences(): boolean
        setEvaluateToError(flag: boolean): void
        isEvaluateToError(): boolean
        setInconsistentFormula(flag: boolean): void
        isInconsistentFormula(): boolean
        setIndicatorColorIndex(l: int): void
        IndicatorColorIndex(): int
        setListDataValidation(flag: boolean): void
        isListDataValidation(): boolean
        setUnlockedFormulaCells(flag: boolean): void
        isUnlockedFormulaCells(): boolean
    }
    export interface Errors {
        getItem(index: int): excel.Error
    }
    export interface Errors {
        getItem(index: int): excel.Error
    }
    export interface FillFormat {
        apply(mFill: any): void
        getType(): int
        setVisible(visible: int): void
        setGradientAngle(value: double): void
        getGradientAngle(): double
        solid(): void
        getBackColor(): excel.ColorFormat
        getForeColor(): excel.ColorFormat
        getGradientStyle(): int
        getPattern(): int
        patterned(pattern: int): void
        presetGradient(style: int, variant: int, presetGradientType: int): void
        getPresetTexture(): int
        presetTextured(texture: int): void
        getTextureName(): string
        getTextureType(): int
        getTransparency(): double
        setTransparency(pTransparency: double): void
        userPicture(pictureFile: string): void
        userTextured(textureFile: string): void
        getVisible(): int
        setTextureTile(textureTile: int): void
        getTextureTile(): int
        oneColorGradient(style: int, variant: int, degree: double): void
        twoColorGradient(style: int, variant: int): void
        getGradientColorType(): int
        getGradientDegree(): double
        transGradientStyleToMS(type: int): int
        changeYzoTypeToMsType(yzoType: int): int
        changeMSTypeToYzoType(msType: int): int
        transGradientStyleToEIO(type: int): int
        getPresetGradientType(): int
        getTextureAlignment(): int
        setTextureAlignment(value: int): void
        getTextureOffsetX(): double
        setTextureOffsetX(x: double): void
        getTextureOffsetY(): double
        setTextureOffsetY(y: double): void
        getTextureVerticalScale(): double
        setTextureVerticalScale(scale: double): void
        getPictureEffects(): office.PictureEffects
        getGradientVariant(): int
        getAngleFromMsStyle(msStyle: int): int
        getTextureHorizontalScale(): double
        setTextureHorizontalScale(value: double): void
        setForeColor(colorFormat: excel.ColorFormat): void
        setBackColor(colorFormat: excel.ColorFormat): void
        transTwoColorGradient(mFill: any, style: int, variant: int): void
        getRotateWithObject(): int
        setRotateWithObject(rotate: int): void
    }
    export interface FillFormat {
        apply(mFill: any): void
        getType(): int
        setVisible(visible: int): void
        setGradientAngle(value: double): void
        getGradientAngle(): double
        solid(): void
        getBackColor(): excel.ColorFormat
        getForeColor(): excel.ColorFormat
        getGradientStyle(): int
        getPattern(): int
        patterned(pattern: int): void
        presetGradient(style: int, variant: int, presetGradientType: int): void
        getPresetTexture(): int
        presetTextured(texture: int): void
        getTextureName(): string
        getTextureType(): int
        getTransparency(): double
        setTransparency(pTransparency: double): void
        userPicture(pictureFile: string): void
        userTextured(textureFile: string): void
        getVisible(): int
        setTextureTile(textureTile: int): void
        getTextureTile(): int
        oneColorGradient(style: int, variant: int, degree: double): void
        twoColorGradient(style: int, variant: int): void
        getGradientColorType(): int
        getGradientDegree(): double
        transGradientStyleToMS(type: int): int
        changeYzoTypeToMsType(yzoType: int): int
        changeMSTypeToYzoType(msType: int): int
        transGradientStyleToEIO(type: int): int
        getPresetGradientType(): int
        getTextureAlignment(): int
        setTextureAlignment(value: int): void
        getTextureOffsetX(): double
        setTextureOffsetX(x: double): void
        getTextureOffsetY(): double
        setTextureOffsetY(y: double): void
        getTextureVerticalScale(): double
        setTextureVerticalScale(scale: double): void
        getPictureEffects(): office.PictureEffects
        getGradientVariant(): int
        getAngleFromMsStyle(msStyle: int): int
        getTextureHorizontalScale(): double
        setTextureHorizontalScale(value: double): void
        setForeColor(colorFormat: excel.ColorFormat): void
        setBackColor(colorFormat: excel.ColorFormat): void
        transTwoColorGradient(mFill: any, style: int, variant: int): void
        getRotateWithObject(): int
        setRotateWithObject(rotate: int): void
    }
    export interface Filter {
        isOn(): boolean
        getCount(): int
        getOperator(): int
        getCriteria1(): any
        setCriteria1(comData1: any): void
        getCriteria2(): any
        setCriteria2(comData2: any): void
        setOperator(op: int): void
        transOperatorToMs(op: int): int
        transOperatorToEIO(op: int): int
    }
    export interface Filter {
        isOn(): boolean
        getCount(): int
        getOperator(): int
        getCriteria1(): any
        setCriteria1(comData1: any): void
        getCriteria2(): any
        setCriteria2(comData2: any): void
        setOperator(op: int): void
        transOperatorToMs(op: int): int
        transOperatorToEIO(op: int): int
    }
    export interface Filters {
        item(index: int): excel.Filter
        getCount(): int
    }
    export interface Filters {
        item(index: int): excel.Filter
        getCount(): int
    }
    export interface Floor {
        getName(): string
        getBorder(): excel.Border
        paste(): void
        select(): any
        getFill(): excel.FillFormat
        clearFormats(): any
        getInterior(): excel.Interior
        getPictureType(): any
        setPictureType(s: any): void
    }
    export interface Floor {
        getName(): string
        getBorder(): excel.Border
        paste(): void
        select(): any
        getFill(): excel.FillFormat
        clearFormats(): any
        getInterior(): excel.Interior
        getPictureType(): any
        setPictureType(s: any): void
    }
    export interface Font {
        getName(): any
        apply(mFontAttr: any): void
        setName(n: any): void
        getSize(): any
        setSize(ff: any): void
        setColor(ll: any): void
        getColor(): any
        setColorIndex(ll: any): void
        getColorIndex(): any
        setShadow(ll: any): void
        setBold(ll: any): void
        getBold(): any
        getFontAttribute(): any
        setFontStyle(s: any): void
        setStrikeThrough(ll: any): void
        setSubscript(ll: any): void
        setSuperscript(ll: any): void
        setItalic(ll: any): void
        getItalic(): any
        getShadow(): any
        getStrikeThrough(): any
        getSubscript(): any
        getSuperscript(): any
        setUnderline(ll: any): void
        getUnderline(): any
        getFontStyle(): any
        setOutlineFont(ll: any): void
        getOutlineFont(): any
        setStrikethrough(ll: any): void
        getStrikethrough(): any
    }
    export interface Font {
        getName(): any
        apply(mFontAttr: any): void
        setName(n: any): void
        getSize(): any
        setSize(ff: any): void
        setColor(ll: any): void
        getColor(): any
        setColorIndex(ll: any): void
        getColorIndex(): any
        setShadow(ll: any): void
        setBold(ll: any): void
        getBold(): any
        getFontAttribute(): any
        setFontStyle(s: any): void
        setStrikeThrough(ll: any): void
        setSubscript(ll: any): void
        setSuperscript(ll: any): void
        setItalic(ll: any): void
        getItalic(): any
        getShadow(): any
        getStrikeThrough(): any
        getSubscript(): any
        getSuperscript(): any
        setUnderline(ll: any): void
        getUnderline(): any
        getFontStyle(): any
        setOutlineFont(ll: any): void
        getOutlineFont(): any
        setStrikethrough(ll: any): void
        getStrikethrough(): any
    }
    export interface FormatColor {
        setColor(rgb2: any): void
        getColor(): any
        setColorIndex(index: int): void
        getColorIndex(): int
    }
    export interface FormatColor {
        setColor(rgb2: any): void
        getColor(): any
        setColorIndex(index: int): void
        getColorIndex(): int
    }
    export interface FormatCondition {
        delete(): void
        setPriority(p: int): void
        getPriority(): int
        getType(): int
        getFont(): excel.Font
        getText(): string
        setText(t: string): void
        getBorders(): excel.Borders
        modify(type: int, operator: any, formula1: any, formula2: any, s: any, operator2: any): void
        getNumberFormat(): any
        setNumberFormat(nf: any): void
        getAppliesTo(): excel.Range
        getDateOperator(date: int): void
        getDateOperator(): int
        getFormula1(): string
        setFormula1(f: string): void
        getFormula2(): string
        setFormula2(f: string): void
        getInterior(): excel.Interior
        getOperator(): int
        setFirstPriority(): void
        setLastPriority(): void
        getStopIfTrue(): boolean
        setStopIfTrue(s: boolean): void
        getTextOperator(): int
        setTextOperator(operator: int): void
        getMFormatCondition(): any
        modifyAppliesToRange(range: excel.Range): void
    }
    export interface FormatCondition {
        delete(): void
        setPriority(p: int): void
        getPriority(): int
        getType(): int
        getFont(): excel.Font
        getText(): string
        setText(t: string): void
        getBorders(): excel.Borders
        modify(type: int, operator: any, formula1: any, formula2: any, s: any, operator2: any): void
        getNumberFormat(): any
        setNumberFormat(nf: any): void
        getAppliesTo(): excel.Range
        getDateOperator(date: int): void
        getDateOperator(): int
        getFormula1(): string
        setFormula1(f: string): void
        getFormula2(): string
        setFormula2(f: string): void
        getInterior(): excel.Interior
        getOperator(): int
        setFirstPriority(): void
        setLastPriority(): void
        getStopIfTrue(): boolean
        setStopIfTrue(s: boolean): void
        getTextOperator(): int
        setTextOperator(operator: int): void
        getMFormatCondition(): any
        modifyAppliesToRange(range: excel.Range): void
    }
    export interface FormatConditions {
        add(type: int, operator: any, formula1: any, formula2: any, string: any, TextOperator: any, DateOperator: any, ScopeType: any): any
        delete(): void
        item(index: any): any
        getCount(): int
        transFormatConditionOperatorToEio(type: int): int
        transFormatConditionTypeToEio(type: int): int
        transFormatConditionOperatorToMs(type: int): int
        transFormatConditionTypeToMs(type: int): int
        transAboveBelowToEio(type: int): int
        transAboveBelowToMs(type: int): int
        addColorScale(colorScaleType: int): any
        addUniqueValues(): any
        addDatabar(): any
        addAboveAverage(): any
        addTop10(): any
        transConditionValueTypesToMS(type: int): int
        transConditionValueTypesToEio(type: int): int
        addIconSetCondition(): any
        transColorScaleTypeToEio(type: int): int
        transDupeUniqueToEio(type: int): int
        transDupeUniqueToMs(type: int): int
        transColorScaleTypeToMS(type: int): int
        transTopBottomTypeToEio(type: int): int
        transTopBottomTypeToMS(type: int): int
        transIconSetIndexToMS(type: int): int
        transIconSetIndexToEIO(type: int): int
    }
    export interface FormatConditions {
        add(type: int, operator: any, formula1: any, formula2: any, string: any, TextOperator: any, DateOperator: any, ScopeType: any): any
        delete(): void
        item(index: any): any
        getCount(): int
        transFormatConditionOperatorToEio(type: int): int
        transFormatConditionTypeToEio(type: int): int
        transFormatConditionOperatorToMs(type: int): int
        transFormatConditionTypeToMs(type: int): int
        transAboveBelowToEio(type: int): int
        transAboveBelowToMs(type: int): int
        addColorScale(colorScaleType: int): any
        addUniqueValues(): any
        addDatabar(): any
        addAboveAverage(): any
        addTop10(): any
        transConditionValueTypesToMS(type: int): int
        transConditionValueTypesToEio(type: int): int
        addIconSetCondition(): any
        transColorScaleTypeToEio(type: int): int
        transDupeUniqueToEio(type: int): int
        transDupeUniqueToMs(type: int): int
        transColorScaleTypeToMS(type: int): int
        transTopBottomTypeToEio(type: int): int
        transTopBottomTypeToMS(type: int): int
        transIconSetIndexToMS(type: int): int
        transIconSetIndexToEIO(type: int): int
    }
    export interface FreeformBuilder {
        addNodes(segmentType: int, editingType: int, X1: double, Y1: double, X2: any, Y2: any, X3: any, Y3: any): void
        convertToShape(): excel.Shape
    }
    export interface FreeformBuilder {
        addNodes(segmentType: int, editingType: int, X1: double, Y1: double, X2: any, Y2: any, X3: any, Y3: any): void
        convertToShape(): excel.Shape
    }
    export const enum Global {
        //clsid=
    }
    export const enum Global {
        //clsid=
    }
    export interface Graphic {
        getFileName(): string
        getWidth(): double
        setWidth(f: double): void
        getHeight(): double
        setHeight(f: double): void
        getBrightness(): double
        setBrightness(f: double): void
        setFileName(s: string): void
        setLockAspectRatio(l: int): void
        getLockAspectRatio(): int
        getColorType(): int
        setColorType(l: int): void
        getContrast(): double
        setContrast(f: double): void
        getCropBottom(): double
        setCropBottom(f: double): void
        getCropLeft(): double
        setCropLeft(f: double): void
        getCropRight(): double
        setCropRight(f: double): void
        getCropTop(): double
        setCropTop(f: double): void
    }
    export interface Graphic {
        getFileName(): string
        getWidth(): double
        setWidth(f: double): void
        getHeight(): double
        setHeight(f: double): void
        getBrightness(): double
        setBrightness(f: double): void
        setFileName(s: string): void
        setLockAspectRatio(l: int): void
        getLockAspectRatio(): int
        getColorType(): int
        setColorType(l: int): void
        getContrast(): double
        setContrast(f: double): void
        getCropBottom(): double
        setCropBottom(f: double): void
        getCropLeft(): double
        setCropLeft(f: double): void
        getCropRight(): double
        setCropRight(f: double): void
        getCropTop(): double
        setCropTop(f: double): void
    }
    export interface Gridlines {
        getName(): string
        delete(): any
        getBorder(): excel.Border
        select(): any
    }
    export interface Gridlines {
        getName(): string
        delete(): any
        getBorder(): excel.Border
        select(): any
    }
    export const enum GroupObject {
        //clsid=
    }
    export const enum GroupObject {
        //clsid=
    }
    export interface GroupShapes {
        item(index: any): excel.Shape
        getShape(shapeName: string): excel.Shape
        getCount(): int
        getRange(index: any): excel.ShapeRange
        getAllShapes(): excel.Shape[]
    }
    export interface GroupShapes {
        item(index: any): excel.Shape
        getShape(shapeName: string): excel.Shape
        getCount(): int
        getRange(index: any): excel.ShapeRange
        getAllShapes(): excel.Shape[]
    }
    export interface HeaderFooter {
        getText(): string
        setText(text: string): void
        getPicture(): excel.Graphic
    }
    export interface HeaderFooter {
        getText(): string
        setText(text: string): void
        getPicture(): excel.Graphic
    }
    export interface HiLoLines {
        getName(): string
        delete(): any
        getBorder(): excel.Border
        select(): any
    }
    export interface HiLoLines {
        getName(): string
        delete(): any
        getBorder(): excel.Border
        select(): any
    }
    export interface HPageBreak {
        getLocation(): excel.Range
        delete(): void
        getType(): int
        setType(type: int): void
        setLocation(range: excel.Range): void
        getExtent(): int
        DragOff(direction: int, regionIndex: int): void
    }
    export interface HPageBreak {
        getLocation(): excel.Range
        delete(): void
        getType(): int
        setType(type: int): void
        setLocation(range: excel.Range): void
        getExtent(): int
        DragOff(direction: int, regionIndex: int): void
    }
    export interface HPageBreaks {
        add(before: excel.Range): excel.HPageBreak
        item(index: int): excel.HPageBreak
        getCount(): int
    }
    export interface HPageBreaks {
        add(before: excel.Range): excel.HPageBreak
        item(index: int): excel.HPageBreak
        getCount(): int
    }
    export interface HTMLProject {
        getState(): int
        open(openKind: int): void
        getHTMLProjectItems(): excel.HTMLProjectItems
        refreshDocument(refresh: boolean): void
        refreshProject(refresh: boolean): void
    }
    export interface HTMLProject {
        getState(): int
        open(openKind: int): void
        getHTMLProjectItems(): excel.HTMLProjectItems
        refreshDocument(refresh: boolean): void
        refreshProject(refresh: boolean): void
    }
    export interface HTMLProjectItem {
        getName(): string
        isOpen(): boolean
        getText(): string
        setText(text: string): void
        loadFromFile(filename: string): void
        saveCopyAs(filename: string): void
        Open(openKind: int): void
    }
    export interface HTMLProjectItem {
        getName(): string
        isOpen(): boolean
        getText(): string
        setText(text: string): void
        loadFromFile(filename: string): void
        saveCopyAs(filename: string): void
        Open(openKind: int): void
    }
    export interface HTMLProjectItems {
        item(index: int): excel.HTMLProjectItem
        getCount(): int
    }
    export interface HTMLProjectItems {
        item(index: int): excel.HTMLProjectItem
        getCount(): int
    }
    export interface Hyperlink {
        getAddress(): string
        getName(): string
        apply(): void
        delete(): void
        getType(): int
        getShape(): excel.Shape
        getRange(): excel.Range
        follow(isNewWindow: boolean, addHistory: any, ExtraInfo: any, method: int, HeaderInfo: string): void
        setTextToDisplay(s: string): void
        setAddress(s: string): void
        setEmailSubject(s: string): void
        getEmailSubject(): string
        setScreenTip(s: string): void
        getScreenTip(): string
        setSubAddress(s: string): void
        getSubAddress(): string
        getTextToDisplay(): string
        addToFavorites(): void
        CreateNewDocument(fileName: string, EditNow: boolean, Overwrite: boolean): void
    }
    export interface Hyperlink {
        getAddress(): string
        getName(): string
        apply(): void
        delete(): void
        getType(): int
        getShape(): excel.Shape
        getRange(): excel.Range
        follow(isNewWindow: boolean, addHistory: any, ExtraInfo: any, method: int, HeaderInfo: string): void
        setTextToDisplay(s: string): void
        setAddress(s: string): void
        setEmailSubject(s: string): void
        getEmailSubject(): string
        setScreenTip(s: string): void
        getScreenTip(): string
        setSubAddress(s: string): void
        getSubAddress(): string
        getTextToDisplay(): string
        addToFavorites(): void
        CreateNewDocument(fileName: string, EditNow: boolean, Overwrite: boolean): void
    }
    export interface Hyperlinks {
        add(anchor: any, address: string, subAddress: any, screenTip: any, textToDisplay: any): any
        item(index: any): excel.Hyperlink
        item(index: int): excel.Hyperlink
        getCount(): int
        Add(anchor: any, address: string, subAddress: string, screenTip: string, textToDisplay: string): excel.Hyperlink
        Delete(): void
    }
    export interface Hyperlinks {
        add(anchor: any, address: string, subAddress: any, screenTip: any, textToDisplay: any): any
        item(index: any): excel.Hyperlink
        item(index: int): excel.Hyperlink
        getCount(): int
        Add(anchor: any, address: string, subAddress: string, screenTip: string, textToDisplay: string): excel.Hyperlink
        Delete(): void
    }
    export interface Icon {
        getIndex(): int
    }
    export interface Icon {
        getIndex(): int
    }
    export interface IconCriteria {
        item(index: int): excel.IconCriterion
        getCount(): int
    }
    export interface IconCriteria {
        item(index: int): excel.IconCriterion
        getCount(): int
    }
    export interface IconCriterion {
        getValue(): any
        setValue(value: any): void
        getType(): int
        getIndex(): int
        setType(type: int): void
        getOperator(): int
        setOperator(type: int): void
        transIconToEIO(index: int): int
        transIconToXLS(index: int): int
        setIcon(id: int): void
        getIcon(): int
    }
    export interface IconCriterion {
        getValue(): any
        setValue(value: any): void
        getType(): int
        getIndex(): int
        setType(type: int): void
        getOperator(): int
        setOperator(type: int): void
        transIconToEIO(index: int): int
        transIconToXLS(index: int): int
        setIcon(id: int): void
        getIcon(): int
    }
    export interface IconSet {
        getID(): int
        item(index: int): excel.Icon
        getCount(): int
    }
    export interface IconSet {
        getID(): int
        item(index: int): excel.Icon
        getCount(): int
    }
    export interface IconSetCondition {
        delete(): void
        setPriority(p: int): void
        getPriority(): int
        getType(): int
        setFormula(f: string): void
        getAppliesTo(): excel.Range
        setFirstPriority(): void
        setLastPriority(): void
        getStopIfTrue(): boolean
        setStopIfTrue(s: boolean): void
        modifyAppliesToRange(range: excel.Range): void
        getIconCriteria(): excel.IconCriteria
        getReverseOrder(): boolean
        setReverseOrder(reverseOrder: boolean): void
        getShowIconOnly(): boolean
        setShowIconOnly(showIconOnly: boolean): void
        getIconSet(): any
        setIconSet(iconSet: any): void
        getFormula(): string
        getPercentileValues(): boolean
        setPercentileValues(percentileValues: boolean): void
    }
    export interface IconSetCondition {
        delete(): void
        setPriority(p: int): void
        getPriority(): int
        getType(): int
        setFormula(f: string): void
        getAppliesTo(): excel.Range
        setFirstPriority(): void
        setLastPriority(): void
        getStopIfTrue(): boolean
        setStopIfTrue(s: boolean): void
        modifyAppliesToRange(range: excel.Range): void
        getIconCriteria(): excel.IconCriteria
        getReverseOrder(): boolean
        setReverseOrder(reverseOrder: boolean): void
        getShowIconOnly(): boolean
        setShowIconOnly(showIconOnly: boolean): void
        getIconSet(): any
        setIconSet(iconSet: any): void
        getFormula(): string
        getPercentileValues(): boolean
        setPercentileValues(percentileValues: boolean): void
    }
    export interface IconSets {
        item(index: int): excel.IconSet
        getCount(): int
    }
    export interface IconSets {
        item(index: int): excel.IconSet
        getCount(): int
    }
    export interface Interior {
        setColor(color: any): void
        getColor(): any
        setColorIndex(obj: any): void
        getColorIndex(): any
        getPattern(): any
        setPattern(pattern: any): void
        getPatternColor(): any
        setPatternColor(c: any): void
        setInvertIfNegative(b: any): void
        isInvertIfNegative(): boolean
        getPatternColorIndex(): any
        setPatternColorIndex(l: any): void
    }
    export interface Interior {
        setColor(color: any): void
        getColor(): any
        setColorIndex(obj: any): void
        getColorIndex(): any
        getPattern(): any
        setPattern(pattern: any): void
        getPatternColor(): any
        setPatternColor(c: any): void
        setInvertIfNegative(b: any): void
        isInvertIfNegative(): boolean
        getPatternColorIndex(): any
        setPatternColorIndex(l: any): void
    }
    export interface LeaderLines {
        delete(): void
        getBorder(): excel.Border
        select(): void
    }
    export interface LeaderLines {
        delete(): void
        getBorder(): excel.Border
        select(): void
    }
    export interface Legend {
        clear(): any
        getName(): string
        delete(): any
        getBorder(): excel.Border
        getFont(): excel.Font
        getWidth(): double
        getLeft(): double
        setLeft(f: double): void
        setTop(f: double): void
        getTop(): double
        setWidth(f: double): void
        getHeight(): double
        setHeight(f: double): void
        select(): any
        isShadow(): boolean
        setPosition(l: int): void
        getPosition(): int
        setShadow(flag: boolean): void
        getFill(): excel.ChartFillFormat
        getInterior(): excel.Interior
        legendEntries(index: any): any
    }
    export interface Legend {
        clear(): any
        getName(): string
        delete(): any
        getBorder(): excel.Border
        getFont(): excel.Font
        getWidth(): double
        getLeft(): double
        setLeft(f: double): void
        setTop(f: double): void
        getTop(): double
        setWidth(f: double): void
        getHeight(): double
        setHeight(f: double): void
        select(): any
        isShadow(): boolean
        setPosition(l: int): void
        getPosition(): int
        setShadow(flag: boolean): void
        getFill(): excel.ChartFillFormat
        getInterior(): excel.Interior
        legendEntries(index: any): any
    }
    export interface LegendEntries {
        item(index: string): excel.LegendEntry
        item(index: int): excel.LegendEntry
        getCount(): int
    }
    export interface LegendEntries {
        item(index: string): excel.LegendEntry
        item(index: int): excel.LegendEntry
        getCount(): int
    }
    export interface LegendEntry {
        delete(): any
        getFont(): excel.Font
        getIndex(): int
        getWidth(): double
        getLeft(): double
        getTop(): double
        getHeight(): double
        select(): any
        getLegendKey(): excel.LegendKey
        isAutoScaleFont(): boolean
        setAutoScaleFont(b: boolean): void
    }
    export interface LegendEntry {
        delete(): any
        getFont(): excel.Font
        getIndex(): int
        getWidth(): double
        getLeft(): double
        getTop(): double
        getHeight(): double
        select(): any
        getLegendKey(): excel.LegendKey
        isAutoScaleFont(): boolean
        setAutoScaleFont(b: boolean): void
    }
    export interface LegendKey {
        delete(): any
        getBorder(): excel.Border
        getWidth(): double
        getLeft(): double
        setLeft(f: double): void
        setTop(f: double): void
        getTop(): double
        setWidth(f: double): void
        getHeight(): double
        setHeight(f: double): void
        select(): void
        isShadow(): boolean
        setShadow(flag: boolean): void
        getFill(): excel.ChartFillFormat
        clearFormats(): any
        getInterior(): excel.Interior
        getPictureType(): int
        setPictureType(l: int): void
        setMarkerSize(l: int): void
        getMarkerSize(): int
        setMarkerStyle(l: int): void
        getMarkerStyle(): int
        setPictureUnit(l: int): void
        getPictureUnit(): int
        setInvertIfNegative(b: boolean): void
        isInvertIfNegative(): boolean
        setMarkerBackgroundColor(l: int): void
        getMarkerBackgroundColor(): int
        setMarkerForegroundColor(l: int): void
        getMarkerForegroundColor(): int
        setMarkerBackgroundColorIndex(l: int): void
        getMarkerBackgroundColorIndex(): int
        setMarkerForegroundColorIndex(l: int): void
        getMarkerForegroundColorIndex(): int
    }
    export interface LegendKey {
        delete(): any
        getBorder(): excel.Border
        getWidth(): double
        getLeft(): double
        setLeft(f: double): void
        setTop(f: double): void
        getTop(): double
        setWidth(f: double): void
        getHeight(): double
        setHeight(f: double): void
        select(): void
        isShadow(): boolean
        setShadow(flag: boolean): void
        getFill(): excel.ChartFillFormat
        clearFormats(): any
        getInterior(): excel.Interior
        getPictureType(): int
        setPictureType(l: int): void
        setMarkerSize(l: int): void
        getMarkerSize(): int
        setMarkerStyle(l: int): void
        getMarkerStyle(): int
        setPictureUnit(l: int): void
        getPictureUnit(): int
        setInvertIfNegative(b: boolean): void
        isInvertIfNegative(): boolean
        setMarkerBackgroundColor(l: int): void
        getMarkerBackgroundColor(): int
        setMarkerForegroundColor(l: int): void
        getMarkerForegroundColor(): int
        setMarkerBackgroundColorIndex(l: int): void
        getMarkerBackgroundColorIndex(): int
        setMarkerForegroundColorIndex(l: int): void
        getMarkerForegroundColorIndex(): int
    }
    export interface Line {
        delete(): void
    }
    export interface Line {
        delete(): void
    }
    export interface LineFormat {
        setVisible(visible: int): void
        getBackColor(): excel.ColorFormat
        getForeColor(): excel.ColorFormat
        getPattern(): int
        getTransparency(): double
        setTransparency(pTransparency: double): void
        getVisible(): int
        setForeColor(colorFormat: excel.ColorFormat): void
        setBackColor(colorFormat: excel.ColorFormat): void
        getDashStyle(): int
        setDashStyle(l: int): void
        setPattern(pattern: int): void
        getWeight(): double
        setWeight(weight: double): void
        getStyle(): int
        setStyle(style: int): void
        getBeginArrowheadLength(): int
        setBeginArrowheadLength(l: int): void
        getBeginArrowheadStyle(): int
        transArrowStyleToMso(type: int): int
        setBeginArrowheadStyle(l: int): void
        transArrowStyleToYzo(type: int): int
        getBeginArrowheadWidth(): int
        setBeginArrowheadWidth(l: int): void
        getEndArrowheadLength(): int
        setEndArrowheadLength(length: int): void
        getEndArrowheadStyle(): int
        setEndArrowheadStyle(l: int): void
        getEndArrowheadWidth(): int
        setEndArrowheadWidth(width: int): void
        transArrowSizeToMSWL(arrowSize: int): number[]
        transArrowWLToEioSize(arrowwidth: int, arrowlength: int): int
    }
    export interface LineFormat {
        setVisible(visible: int): void
        getBackColor(): excel.ColorFormat
        getForeColor(): excel.ColorFormat
        getPattern(): int
        getTransparency(): double
        setTransparency(pTransparency: double): void
        getVisible(): int
        setForeColor(colorFormat: excel.ColorFormat): void
        setBackColor(colorFormat: excel.ColorFormat): void
        getDashStyle(): int
        setDashStyle(l: int): void
        setPattern(pattern: int): void
        getWeight(): double
        setWeight(weight: double): void
        getStyle(): int
        setStyle(style: int): void
        getBeginArrowheadLength(): int
        setBeginArrowheadLength(l: int): void
        getBeginArrowheadStyle(): int
        transArrowStyleToMso(type: int): int
        setBeginArrowheadStyle(l: int): void
        transArrowStyleToYzo(type: int): int
        getBeginArrowheadWidth(): int
        setBeginArrowheadWidth(l: int): void
        getEndArrowheadLength(): int
        setEndArrowheadLength(length: int): void
        getEndArrowheadStyle(): int
        setEndArrowheadStyle(l: int): void
        getEndArrowheadWidth(): int
        setEndArrowheadWidth(width: int): void
        transArrowSizeToMSWL(arrowSize: int): number[]
        transArrowWLToEioSize(arrowwidth: int, arrowlength: int): int
    }
    export interface LinkFormat {
        update(): void
        isLocked(): boolean
        isAutoUpdate(): boolean
        setLocked(b: boolean): void
    }
    export interface LinkFormat {
        update(): void
        isLocked(): boolean
        isAutoUpdate(): boolean
        setLocked(b: boolean): void
    }
    export interface ListColumn {
        getName(): string
        delete(): void
        setName(s: string): void
        getIndex(): int
        getTotal(): excel.Range
        getRange(): excel.Range
        getDataBodyRange(): excel.Range
        getXPath(): excel.XPath
        getListDataFormat(): excel.ListDataFormat
        getSharePointFormula(): string
        setTotalsCalculation(totalsCalculation: int): void
        getTotalsCalculation(): int
    }
    export interface ListColumn {
        getName(): string
        delete(): void
        setName(s: string): void
        getIndex(): int
        getTotal(): excel.Range
        getRange(): excel.Range
        getDataBodyRange(): excel.Range
        getXPath(): excel.XPath
        getListDataFormat(): excel.ListDataFormat
        getSharePointFormula(): string
        setTotalsCalculation(totalsCalculation: int): void
        getTotalsCalculation(): int
    }
    export interface ListColumns {
        add(position: any): excel.ListColumn
        getCount(): int
        getItem(index: any): excel.ListColumn
    }
    export interface ListColumns {
        add(position: any): excel.ListColumn
        getCount(): int
        getItem(index: any): excel.ListColumn
    }
    export interface ListDataFormat {
        getType(): int
        getDefaultValue(): any
        isReadOnly(): boolean
        isPercent(): boolean
        isAllowFillIn(): boolean
        getChoices(): any
        getDecimalPlaces(): int
        getMaxCharacters(): int
        getMaxNumber(): any
        getMinNumber(): any
        isRequired(): boolean
        getlcid(): int
    }
    export interface ListDataFormat {
        getType(): int
        getDefaultValue(): any
        isReadOnly(): boolean
        isPercent(): boolean
        isAllowFillIn(): boolean
        getChoices(): any
        getDecimalPlaces(): int
        getMaxCharacters(): int
        getMaxNumber(): any
        getMinNumber(): any
        isRequired(): boolean
        getlcid(): int
    }
    export interface ListObject {
        getName(): string
        delete(): void
        setName(s: string): void
        resize(range: excel.Range): void
        resize(rowSize: int, columnSize: int): excel.Range
        getDisplayName(): string
        unlink(): void
        getComment(): string
        setComment(comment: string): void
        isActive(): boolean
        refresh(): void
        getRange(): excel.Range
        setAlternativeText(text: string): void
        getAlternativeText(): string
        getDataBodyRange(): excel.Range
        getAutoFilter(): excel.AutoFilter
        setDisplayName(name: string): void
        exportToVisio(): void
        getListColumns(): excel.ListColumns
        getListRows(): excel.ListRows
        getQueryTable(): excel.QueryTable
        getSharePointURL(): string
        isShowAutoFilter(): boolean
        setShowHeaders(b: boolean): void
        isShowHeaders(): boolean
        setShowTotals(b: boolean): void
        isShowTotals(): boolean
        getSourceType(): int
        setSummary(text: string): void
        getSummary(): string
        UpdateChanges(iConflictType: int): void
        getXmlMap(): excel.XmlMap
        Publish(target: string[], linkSource: boolean): void
        getSort(): excel.Sort
        unlist(): void
        isDisplayRightToLeft(): boolean
        getHeaderRowRange(): excel.Range
        getInsertRowRange(): excel.Range
        setShowAutoFilter(b: boolean): void
        getTotalsRowRange(): excel.Range
        setShowTableStyleColumnStripes(b: boolean): void
        isShowTableStyleColumnStripes(): boolean
        setShowTableStyleFirstColumn(b: boolean): void
        isShowTableStyleFirstColumn(): boolean
        setShowTableStyleLastColumn(b: boolean): void
        isShowTableStyleLastColumn(): boolean
        setShowTableStyleRowStripes(b: boolean): void
        isShowTableStyleRowStripes(): boolean
    }
    export interface ListObject {
        getName(): string
        delete(): void
        setName(s: string): void
        resize(range: excel.Range): void
        resize(rowSize: int, columnSize: int): excel.Range
        getDisplayName(): string
        unlink(): void
        getComment(): string
        setComment(comment: string): void
        isActive(): boolean
        refresh(): void
        getRange(): excel.Range
        setAlternativeText(text: string): void
        getAlternativeText(): string
        getDataBodyRange(): excel.Range
        getAutoFilter(): excel.AutoFilter
        setDisplayName(name: string): void
        exportToVisio(): void
        getListColumns(): excel.ListColumns
        getListRows(): excel.ListRows
        getQueryTable(): excel.QueryTable
        getSharePointURL(): string
        isShowAutoFilter(): boolean
        setShowHeaders(b: boolean): void
        isShowHeaders(): boolean
        setShowTotals(b: boolean): void
        isShowTotals(): boolean
        getSourceType(): int
        setSummary(text: string): void
        getSummary(): string
        UpdateChanges(iConflictType: int): void
        getXmlMap(): excel.XmlMap
        Publish(target: string[], linkSource: boolean): void
        getSort(): excel.Sort
        unlist(): void
        isDisplayRightToLeft(): boolean
        getHeaderRowRange(): excel.Range
        getInsertRowRange(): excel.Range
        setShowAutoFilter(b: boolean): void
        getTotalsRowRange(): excel.Range
        setShowTableStyleColumnStripes(b: boolean): void
        isShowTableStyleColumnStripes(): boolean
        setShowTableStyleFirstColumn(b: boolean): void
        isShowTableStyleFirstColumn(): boolean
        setShowTableStyleLastColumn(b: boolean): void
        isShowTableStyleLastColumn(): boolean
        setShowTableStyleRowStripes(b: boolean): void
        isShowTableStyleRowStripes(): boolean
    }
    export interface ListObjects {
        add(sourceType: int, source: any, linkSource: any, hasHeaders: any, destination: any): excel.ListObject
        item(index: any): excel.ListObject
        getCount(): int
    }
    export interface ListObjects {
        add(sourceType: int, source: any, linkSource: any, hasHeaders: any, destination: any): excel.ListObject
        item(index: any): excel.ListObject
        getCount(): int
    }
    export interface ListRow {
        delete(): void
        getIndex(): int
        getRange(): excel.Range
        isInvalidData(): boolean
    }
    export interface ListRow {
        delete(): void
        getIndex(): int
        getRange(): excel.Range
        isInvalidData(): boolean
    }
    export interface ListRows {
        add(position: any, alwaysInsert: any): excel.ListRow
        getCount(): int
        getItem(index: any): excel.ListRow
    }
    export interface ListRows {
        add(position: any, alwaysInsert: any): excel.ListRow
        getCount(): int
        getItem(index: any): excel.ListRow
    }
    export const enum MajorGridlines {
        //clsid=
    }
    export const enum MajorGridlines {
        //clsid=
    }
    export interface Menu {
        delete(): void
        getIndex(): int
        getCaption(): string
        setEnabled(enabled: boolean): void
        isEnabled(): boolean
        getMenuItems(): excel.MenuItems
    }
    export interface Menu {
        delete(): void
        getIndex(): int
        getCaption(): string
        setEnabled(enabled: boolean): void
        isEnabled(): boolean
        getMenuItems(): excel.MenuItems
    }
    export interface MenuBar {
        delete(): void
        reset(): void
        getIndex(): int
        activate(): void
        setCaption(caption: string): void
        getCaption(): string
        getBuiltIn(): boolean
        getMenus(): excel.Menus
    }
    export interface MenuBar {
        delete(): void
        reset(): void
        getIndex(): int
        activate(): void
        setCaption(caption: string): void
        getCaption(): string
        getBuiltIn(): boolean
        getMenus(): excel.Menus
    }
    export interface MenuBars {
        add(name: string): excel.MenuBar
        item(obj: any): excel.MenuBar
        getCount(): int
    }
    export interface MenuBars {
        add(name: string): excel.MenuBar
        item(obj: any): excel.MenuBar
        getCount(): int
    }
    export interface MenuItem {
        delete(): void
        getIndex(): int
        getCaption(): string
        getStatusBar(): string
        setStatusBar(statusBar: string): void
        setEnabled(enabled: boolean): void
        isEnabled(): boolean
        isChecked(): boolean
        setChecked(checked: boolean): void
        getOnAction(): string
        setOnAction(onAction: string): void
        setHelpContextID(helpContextID: int): void
        getHelpContextID(): int
        setHelpFile(helpFile: string): void
        getHelpFile(): string
    }
    export interface MenuItem {
        delete(): void
        getIndex(): int
        getCaption(): string
        getStatusBar(): string
        setStatusBar(statusBar: string): void
        setEnabled(enabled: boolean): void
        isEnabled(): boolean
        isChecked(): boolean
        setChecked(checked: boolean): void
        getOnAction(): string
        setOnAction(onAction: string): void
        setHelpContextID(helpContextID: int): void
        getHelpContextID(): int
        setHelpFile(helpFile: string): void
        getHelpFile(): string
    }
    export interface MenuItems {
        add(caption: string, onAction: string, shortcutKey: any, before: any, restore: any, statusBar: any, helpFile: any, helpContextID: any): excel.MenuItem
        item(obj: any): any
        getCount(): int
        addMenu(caption: string, before: any, restore: any): excel.Menu
    }
    export interface MenuItems {
        add(caption: string, onAction: string, shortcutKey: any, before: any, restore: any, statusBar: any, helpFile: any, helpContextID: any): excel.MenuItem
        item(obj: any): any
        getCount(): int
        addMenu(caption: string, before: any, restore: any): excel.Menu
    }
    export interface Menus {
        add(caption: string, before: any, restore: any): excel.Menu
        item(obj: any): excel.Menu
        getCount(): int
    }
    export interface Menus {
        add(caption: string, before: any, restore: any): excel.Menu
        item(obj: any): excel.Menu
        getCount(): int
    }
    export const enum MinorGridlines {
        //clsid=
    }
    export const enum MinorGridlines {
        //clsid=
    }
    export interface Name {
        getName(): string
        getValue(): string
        delete(): void
        setName(name: string): void
        setValue(value: string): void
        getIndex(): int
        getComment(): string
        setComment(comment: string): void
        isVisible(): boolean
        setVisible(visible: boolean): void
        getCategory(): string
        getNameLocal(): string
        setNameLocal(nameLocal: string): void
        setCategory(category: string): void
        getRefersTo(): any
        getCategoryLocal(): string
        setCategoryLocal(categoryLocal: string): void
        getMacroType(): int
        setMacroType(macroType: int): void
        setRefersTo(refersTo: any): void
        getRefersToLocal(): any
        setRefersToLocal(refersToLocal: any): void
        getRefersToR1C1(): any
        setRefersToR1C1(refersToR1C1: any): void
        getRefersToRange(): excel.Range
        getShortcutKey(): string
        setShortcutKey(shortcutKey: string): void
        getRefersToR1C1Local(): any
        setRefersToR1C1Local(refersToR1C1Local: any): void
    }
    export interface Name {
        getName(): string
        getValue(): string
        delete(): void
        setName(name: string): void
        setValue(value: string): void
        getIndex(): int
        getComment(): string
        setComment(comment: string): void
        isVisible(): boolean
        setVisible(visible: boolean): void
        getCategory(): string
        getNameLocal(): string
        setNameLocal(nameLocal: string): void
        setCategory(category: string): void
        getRefersTo(): any
        getCategoryLocal(): string
        setCategoryLocal(categoryLocal: string): void
        getMacroType(): int
        setMacroType(macroType: int): void
        setRefersTo(refersTo: any): void
        getRefersToLocal(): any
        setRefersToLocal(refersToLocal: any): void
        getRefersToR1C1(): any
        setRefersToR1C1(refersToR1C1: any): void
        getRefersToRange(): excel.Range
        getShortcutKey(): string
        setShortcutKey(shortcutKey: string): void
        getRefersToR1C1Local(): any
        setRefersToR1C1Local(refersToR1C1Local: any): void
    }
    export interface Names {
        add(name: any, refersTo: any, visible: any, macroType: any, shortcutKey: any, category: any, nameLocal: any, refersToLocal: any, categoryLocal: any, refersToR1C1: any, refersToR1C1Local: any): excel.Name
        item(index: any, indexLocal: any, refersTo: any): excel.Name
        getCount(): int
    }
    export interface Names {
        add(name: any, refersTo: any, visible: any, macroType: any, shortcutKey: any, category: any, nameLocal: any, refersToLocal: any, categoryLocal: any, refersToR1C1: any, refersToR1C1Local: any): excel.Name
        item(index: any, indexLocal: any, refersTo: any): excel.Name
        getCount(): int
    }
    export interface NegativeBarFormat {
        setColor(Color: any): void
        getColor(): any
        setBorderColor(borderColor: any): void
        getColorType(): int
        setColorType(same: int): void
        getMNegativeBarFormat(): any
        getBorderColorType(): int
        setBorderColorType(same: int): void
        getBorderColor(): any
    }
    export interface NegativeBarFormat {
        setColor(Color: any): void
        getColor(): any
        setBorderColor(borderColor: any): void
        getColorType(): int
        setColorType(same: int): void
        getMNegativeBarFormat(): any
        getBorderColorType(): int
        setBorderColorType(same: int): void
        getBorderColor(): any
    }
    export interface ODBCError {
        sqlState(): string
        getErrorString(): string
    }
    export interface ODBCError {
        sqlState(): string
        getErrorString(): string
    }
    export interface ODBCErrors {
        item(index: int): excel.ODBCError
        getCount(): int
    }
    export interface ODBCErrors {
        item(index: int): excel.ODBCError
        getCount(): int
    }
    export interface OLEDBError {
        getNumber(): int
        getStage(): int
        getErrorString(): string
        getNative(): int
        getSqlState(): string
    }
    export interface OLEDBError {
        getNumber(): int
        getStage(): int
        getErrorString(): string
        getNative(): int
        getSqlState(): string
    }
    export interface OLEDBErrors {
        item(index: int): excel.OLEDBError
        getCount(): int
    }
    export interface OLEDBErrors {
        item(index: int): excel.OLEDBError
        getCount(): int
    }
    export interface OLEFormat {
        getObject(): any
        activate(): void
        getProgID(): string
        verb(verb: int): void
    }
    export interface OLEFormat {
        getObject(): any
        activate(): void
        getProgID(): string
        verb(verb: int): void
    }
    export interface OLEObject {
        update(): any
        getObject(): any
        getName(): string
        delete(): any
        setName(name: string): void
        copy(): any
        isLocked(): boolean
        duplicate(): any
        getBorder(): excel.Border
        getIndex(): int
        getWidth(): double
        getLeft(): double
        setLeft(left: double): void
        setTop(top: double): void
        getTop(): double
        setWidth(width: double): void
        activate(): any
        getHeight(): double
        setHeight(height: double): void
        isVisible(): boolean
        setVisible(visible: boolean): void
        getNativeHandle(): int
        setNativeHandle(handle: int): void
        cut(): any
        setEnabled(enabled: boolean): void
        isEnabled(): boolean
        select(replace: any): any
        isShadow(): boolean
        setShadow(shadow: boolean): void
        getShapeRange(): excel.ShapeRange
        isAutoUpdate(): boolean
        setAutoUpdate(autoUpdate: boolean): void
        getSourceName(): string
        setLocked(locked: boolean): void
        getProgID(): string
        getZOrder(): int
        getInterior(): excel.Interior
        copyPicture(appearance: int, format: int): any
        getTopLeftCell(): excel.Range
        setPlacement(placement: any): void
        getPlacement(): any
        setPrintObject(printObject: boolean): void
        isPrintObject(): boolean
        bringToFront(): any
        sendToBack(): any
        getLinkedCell(): string
        setLinkedCell(linkedCell: string): void
        getListFillRange(): string
        setListFillRange(listFillRange: string): void
        verb(verb: int): any
        addOLEObjectListener(listener: excel.event.OLEObjectListener): void
        removeOLEObjectListener(listener: excel.event.OLEObjectListener): void
        getBottomRightCell(): excel.Range
        isAutoLoad(): boolean
        setAutoLoad(autoLoad: boolean): void
        getOLEType(): any
        setSourceName(sourceName: string): void
    }
    export interface OLEObject {
        update(): any
        getObject(): any
        getName(): string
        delete(): any
        setName(name: string): void
        copy(): any
        isLocked(): boolean
        duplicate(): any
        getBorder(): excel.Border
        getIndex(): int
        getWidth(): double
        getLeft(): double
        setLeft(left: double): void
        setTop(top: double): void
        getTop(): double
        setWidth(width: double): void
        activate(): any
        getHeight(): double
        setHeight(height: double): void
        isVisible(): boolean
        setVisible(visible: boolean): void
        getNativeHandle(): int
        setNativeHandle(handle: int): void
        cut(): any
        setEnabled(enabled: boolean): void
        isEnabled(): boolean
        select(replace: any): any
        isShadow(): boolean
        setShadow(shadow: boolean): void
        getShapeRange(): excel.ShapeRange
        isAutoUpdate(): boolean
        setAutoUpdate(autoUpdate: boolean): void
        getSourceName(): string
        setLocked(locked: boolean): void
        getProgID(): string
        getZOrder(): int
        getInterior(): excel.Interior
        copyPicture(appearance: int, format: int): any
        getTopLeftCell(): excel.Range
        setPlacement(placement: any): void
        getPlacement(): any
        setPrintObject(printObject: boolean): void
        isPrintObject(): boolean
        bringToFront(): any
        sendToBack(): any
        getLinkedCell(): string
        setLinkedCell(linkedCell: string): void
        getListFillRange(): string
        setListFillRange(listFillRange: string): void
        verb(verb: int): any
        addOLEObjectListener(listener: excel.event.OLEObjectListener): void
        removeOLEObjectListener(listener: excel.event.OLEObjectListener): void
        getBottomRightCell(): excel.Range
        isAutoLoad(): boolean
        setAutoLoad(autoLoad: boolean): void
        getOLEType(): any
        setSourceName(sourceName: string): void
    }
    export interface OLEObjects {
        delete(): any
        copy(): any
        isLocked(): boolean
        duplicate(): any
        getBorder(): excel.Border
        item(index: any): any
        getCount(): int
        getWidth(): double
        getLeft(): double
        setLeft(left: double): void
        setTop(top: double): void
        getTop(): double
        setWidth(width: double): void
        getHeight(): double
        setHeight(height: double): void
        isVisible(): boolean
        setVisible(visible: boolean): void
        cut(): any
        setEnabled(enabled: boolean): void
        isEnabled(): boolean
        select(replace: any): any
        isShadow(): boolean
        setShadow(shadow: boolean): void
        Add(classType: string, filename: string, link: boolean, displayAsIcon: boolean, iconFileName: string, iconIndex: int, iconLabel: string, left: double, top: double, width: double, height: double): excel.OLEObject
        getShapeRange(): excel.ShapeRange
        getSourceName(): string
        setLocked(locked: boolean): void
        getZOrder(): int
        getInterior(): excel.Interior
        copyPicture(appearance: int, format: int): any
        setPlacement(placement: any): void
        getPlacement(): any
        setPrintObject(printObject: boolean): void
        isPrintObject(): boolean
        bringToFront(): any
        sendToBack(): any
        isAutoLoad(): boolean
        setAutoLoad(autoLoad: boolean): void
        setSourceName(sourceName: string): void
    }
    export interface OLEObjects {
        delete(): any
        copy(): any
        isLocked(): boolean
        duplicate(): any
        getBorder(): excel.Border
        item(index: any): any
        getCount(): int
        getWidth(): double
        getLeft(): double
        setLeft(left: double): void
        setTop(top: double): void
        getTop(): double
        setWidth(width: double): void
        getHeight(): double
        setHeight(height: double): void
        isVisible(): boolean
        setVisible(visible: boolean): void
        cut(): any
        setEnabled(enabled: boolean): void
        isEnabled(): boolean
        select(replace: any): any
        isShadow(): boolean
        setShadow(shadow: boolean): void
        Add(classType: string, filename: string, link: boolean, displayAsIcon: boolean, iconFileName: string, iconIndex: int, iconLabel: string, left: double, top: double, width: double, height: double): excel.OLEObject
        getShapeRange(): excel.ShapeRange
        getSourceName(): string
        setLocked(locked: boolean): void
        getZOrder(): int
        getInterior(): excel.Interior
        copyPicture(appearance: int, format: int): any
        setPlacement(placement: any): void
        getPlacement(): any
        setPrintObject(printObject: boolean): void
        isPrintObject(): boolean
        bringToFront(): any
        sendToBack(): any
        isAutoLoad(): boolean
        setAutoLoad(autoLoad: boolean): void
        setSourceName(sourceName: string): void
    }
    export const enum Options {

    }
    export const enum Options {

    }
    export interface Outline {
        isAutomaticStyles(): boolean
        setAutomaticStyles(automaticStyles: boolean): void
        getSummaryRow(): int
        setSummaryRow(summaryRow: int): void
        getSummaryColumn(): int
        setSummaryColumn(summaryColumn: int): void
        showLevels(rowLevels: int, columnLevels: int): void
    }
    export interface Outline {
        isAutomaticStyles(): boolean
        setAutomaticStyles(automaticStyles: boolean): void
        getSummaryRow(): int
        setSummaryRow(summaryRow: int): void
        getSummaryColumn(): int
        setSummaryColumn(summaryColumn: int): void
        showLevels(rowLevels: int, columnLevels: int): void
    }
    export const enum Oval {
        //clsid=
    }
    export const enum Oval {
        //clsid=
    }
    export interface Page {
        getCenterFooter(): string
        getCenterHeader(): string
        getLeftFooter(): string
        getLeftHeader(): string
        getRightFooter(): string
        getRightHeader(): string
    }
    export interface Page {
        getCenterFooter(): string
        getCenterHeader(): string
        getLeftFooter(): string
        getLeftHeader(): string
        getRightFooter(): string
        getRightHeader(): string
    }
    export interface Pages {
        item(index: int): excel.Page
        getCount(): int
    }
    export interface Pages {
        item(index: int): excel.Page
        getCount(): int
    }
    export interface PageSetup {
        getOrientation(): int
        setPrintComments(printComments: int): void
        setOrientation(orientation: int): void
        setDifferentFirstPageHeaderFooter(b: boolean): void
        setOddAndEvenPagesHeaderFooter(b: boolean): void
        getPages(): excel.Pages
        getZoom(): any
        setTopMargin(topMargin: double): void
        getTopMargin(): double
        getBottomMargin(): double
        setBottomMargin(bottomMargin: double): void
        getLeftMargin(): double
        setLeftMargin(leftMargin: double): void
        getPaperSize(): int
        setPaperSize(paperSize: int): void
        getRightMargin(): double
        setRightMargin(rightMargin: double): void
        isDraft(): boolean
        setDraft(draft: boolean): void
        setZoom(zoom: any): void
        getOrder(): int
        setOrder(order: int): void
        getCenterFooterPicture(): excel.Graphic
        getCenterHeaderPicture(): excel.Graphic
        isCenterHorizontally(): boolean
        setCenterHorizontally(centerHorizontally: boolean): void
        isCenterVertically(): boolean
        setCenterVertically(centerVertically: boolean): void
        getFirstPageNumber(): int
        setFirstPageNumber(firstPageNumber: int): void
        getFitToPagesTall(): any
        setFitToPagesTall(fitToPagesTall: any): void
        getFitToPagesWide(): any
        setFitToPagesWide(fitToPagesWide: any): void
        getLeftFooterPicture(): excel.Graphic
        getLeftHeaderPicture(): excel.Graphic
        setPrintGridlines(printGridlines: boolean): void
        getPrintTitleColumns(): string
        setPrintTitleColumns(printTitleColumns: string): void
        getPrintTitleRows(): string
        setPrintTitleRows(printTitleRows: string): void
        getRightFooterPicture(): excel.Graphic
        getRightHeaderPicture(): excel.Graphic
        isDifferentFirstPageHeaderFooter(): boolean
        isOddAndEvenPagesHeaderFooter(): boolean
        getCenterFooter(): string
        getCenterHeader(): string
        getLeftFooter(): string
        getLeftHeader(): string
        getRightFooter(): string
        getRightHeader(): string
        isBlackAndWhite(): boolean
        setBlackAndWhite(blackAndWhite: boolean): void
        setCenterFooter(centerFooter: string): void
        setCenterHeader(centerHeader: string): void
        getChartSize(): int
        setChartSize(chartSize: int): void
        getFooterMargin(): double
        setFooterMargin(footerMargin: double): void
        getHeaderMargin(): double
        setHeaderMargin(headerMargin: double): void
        setLeftFooter(leftFooter: string): void
        setLeftHeader(leftHeader: string): void
        getPrintArea(): string
        setPrintArea(printArea: string): void
        getPrintComments(): int
        getPrintErrors(): int
        setPrintErrors(printErrors: int): void
        isPrintGridlines(): boolean
        isPrintHeadings(): boolean
        setPrintHeadings(printHeadings: boolean): void
        isPrintNotes(): boolean
        setPrintNotes(printNotes: boolean): void
        getPrintQuality(index: any): any
        setPrintQuality(index: any, printQuality: any): void
        setRightFooter(rightFooter: string): void
        setRightHeader(rightHeader: string): void
    }
    export interface PageSetup {
        getOrientation(): int
        setPrintComments(printComments: int): void
        setOrientation(orientation: int): void
        setDifferentFirstPageHeaderFooter(b: boolean): void
        setOddAndEvenPagesHeaderFooter(b: boolean): void
        getPages(): excel.Pages
        getZoom(): any
        setTopMargin(topMargin: double): void
        getTopMargin(): double
        getBottomMargin(): double
        setBottomMargin(bottomMargin: double): void
        getLeftMargin(): double
        setLeftMargin(leftMargin: double): void
        getPaperSize(): int
        setPaperSize(paperSize: int): void
        getRightMargin(): double
        setRightMargin(rightMargin: double): void
        isDraft(): boolean
        setDraft(draft: boolean): void
        setZoom(zoom: any): void
        getOrder(): int
        setOrder(order: int): void
        getCenterFooterPicture(): excel.Graphic
        getCenterHeaderPicture(): excel.Graphic
        isCenterHorizontally(): boolean
        setCenterHorizontally(centerHorizontally: boolean): void
        isCenterVertically(): boolean
        setCenterVertically(centerVertically: boolean): void
        getFirstPageNumber(): int
        setFirstPageNumber(firstPageNumber: int): void
        getFitToPagesTall(): any
        setFitToPagesTall(fitToPagesTall: any): void
        getFitToPagesWide(): any
        setFitToPagesWide(fitToPagesWide: any): void
        getLeftFooterPicture(): excel.Graphic
        getLeftHeaderPicture(): excel.Graphic
        setPrintGridlines(printGridlines: boolean): void
        getPrintTitleColumns(): string
        setPrintTitleColumns(printTitleColumns: string): void
        getPrintTitleRows(): string
        setPrintTitleRows(printTitleRows: string): void
        getRightFooterPicture(): excel.Graphic
        getRightHeaderPicture(): excel.Graphic
        isDifferentFirstPageHeaderFooter(): boolean
        isOddAndEvenPagesHeaderFooter(): boolean
        getCenterFooter(): string
        getCenterHeader(): string
        getLeftFooter(): string
        getLeftHeader(): string
        getRightFooter(): string
        getRightHeader(): string
        isBlackAndWhite(): boolean
        setBlackAndWhite(blackAndWhite: boolean): void
        setCenterFooter(centerFooter: string): void
        setCenterHeader(centerHeader: string): void
        getChartSize(): int
        setChartSize(chartSize: int): void
        getFooterMargin(): double
        setFooterMargin(footerMargin: double): void
        getHeaderMargin(): double
        setHeaderMargin(headerMargin: double): void
        setLeftFooter(leftFooter: string): void
        setLeftHeader(leftHeader: string): void
        getPrintArea(): string
        setPrintArea(printArea: string): void
        getPrintComments(): int
        getPrintErrors(): int
        setPrintErrors(printErrors: int): void
        isPrintGridlines(): boolean
        isPrintHeadings(): boolean
        setPrintHeadings(printHeadings: boolean): void
        isPrintNotes(): boolean
        setPrintNotes(printNotes: boolean): void
        getPrintQuality(index: any): any
        setPrintQuality(index: any, printQuality: any): void
        setRightFooter(rightFooter: string): void
        setRightHeader(rightHeader: string): void
    }
    export interface Pane {
        getIndex(): int
        activate(): boolean
        largeScroll(down: any, up: any, toRight: any, toLeft: any): any
        smallScroll(down: any, up: any, toRight: any, toLeft: any): any
        scrollIntoView(left: int, top: int, width: int, height: int, start: any): void
        getVisibleRange(): excel.Range
        getScrollColumn(): int
        setScrollColumn(scrollColumn: int): void
        getScrollRow(): int
        setScrollRow(scrollRow: int): void
    }
    export interface Pane {
        getIndex(): int
        activate(): boolean
        largeScroll(down: any, up: any, toRight: any, toLeft: any): any
        smallScroll(down: any, up: any, toRight: any, toLeft: any): any
        scrollIntoView(left: int, top: int, width: int, height: int, start: any): void
        getVisibleRange(): excel.Range
        getScrollColumn(): int
        setScrollColumn(scrollColumn: int): void
        getScrollRow(): int
        setScrollRow(scrollRow: int): void
    }
    export interface Panes {
        item(index: int): excel.Pane
        getCount(): int
    }
    export interface Panes {
        item(index: int): excel.Pane
        getCount(): int
    }
    export interface Parameter {
        getName(): string
        getValue(): any
        setName(name: string): void
        getType(): int
        setParam(type: int, value: any): void
        isRefreshOnChange(): boolean
        setRefreshOnChange(refreshOnChange: boolean): void
        getDataType(): int
        setDataType(dataType: int): void
        getPromptString(): string
        getSourceRange(): excel.Range
    }
    export interface Parameter {
        getName(): string
        getValue(): any
        setName(name: string): void
        getType(): int
        setParam(type: int, value: any): void
        isRefreshOnChange(): boolean
        setRefreshOnChange(refreshOnChange: boolean): void
        getDataType(): int
        setDataType(dataType: int): void
        getPromptString(): string
        getSourceRange(): excel.Range
    }
    export interface Parameters {
        add(name: string, iDataType: int): excel.Parameter
        delete(): void
        item(index: any): excel.Parameter
        getCount(): int
    }
    export interface Parameters {
        add(name: string, iDataType: int): excel.Parameter
        delete(): void
        item(index: any): excel.Parameter
        getCount(): int
    }
    export interface Phonetic {
        getFont(): excel.Font
        isVisible(): boolean
        setVisible(visible: boolean): void
        getText(): string
        setText(text: string): void
        getAlignment(): int
        setAlignment(alignment: int): void
        getCharacterType(): int
        setCharacterType(characterType: int): void
    }
    export interface Phonetic {
        getFont(): excel.Font
        isVisible(): boolean
        setVisible(visible: boolean): void
        getText(): string
        setText(text: string): void
        getAlignment(): int
        setAlignment(alignment: int): void
        getCharacterType(): int
        setCharacterType(characterType: int): void
    }
    export interface Phonetics {
        add(start: int, length: int, text: string): void
        getLength(): int
        delete(): void
        getFont(): excel.Font
        getCount(): int
        getItem(index: int): excel.Phonetic
        isVisible(): boolean
        setVisible(visible: boolean): void
        getText(): string
        setText(text: string): void
        getStart(): int
        getAlignment(): int
        setAlignment(alignment: int): void
        getCharacterType(): int
        setCharacterType(characterType: int): void
    }
    export interface Phonetics {
        add(start: int, length: int, text: string): void
        getLength(): int
        delete(): void
        getFont(): excel.Font
        getCount(): int
        getItem(index: int): excel.Phonetic
        isVisible(): boolean
        setVisible(visible: boolean): void
        getText(): string
        setText(text: string): void
        getStart(): int
        getAlignment(): int
        setAlignment(alignment: int): void
        getCharacterType(): int
        setCharacterType(characterType: int): void
    }
    export const enum Picture {
        //clsid=
    }
    export const enum Picture {
        //clsid=
    }
    export interface PictureFormat {
        getBrightness(): double
        setBrightness(brightness: double): void
        getTransparencyColor(): int
        setTransparencyColor(transparencyColor: int): void
        getTransparentBackground(): int
        setTransparentBackground(transparentBackground: int): void
        incrementBrightness(increment: float): void
        incrementContrast(increment: float): void
        getColorType(): int
        setColorType(colorType: int): void
        getContrast(): double
        setContrast(contrast: double): void
        getCropBottom(): double
        setCropBottom(cropBottom: double): void
        getCropLeft(): double
        setCropLeft(cropLeft: double): void
        getCropRight(): double
        setCropRight(cropRight: double): void
        getCropTop(): double
        setCropTop(cropTop: double): void
    }
    export interface PictureFormat {
        getBrightness(): double
        setBrightness(brightness: double): void
        getTransparencyColor(): int
        setTransparencyColor(transparencyColor: int): void
        getTransparentBackground(): int
        setTransparentBackground(transparentBackground: int): void
        incrementBrightness(increment: float): void
        incrementContrast(increment: float): void
        getColorType(): int
        setColorType(colorType: int): void
        getContrast(): double
        setContrast(contrast: double): void
        getCropBottom(): double
        setCropBottom(cropBottom: double): void
        getCropLeft(): double
        setCropLeft(cropLeft: double): void
        getCropRight(): double
        setCropRight(cropRight: double): void
        getCropTop(): double
        setCropTop(cropTop: double): void
    }
    export interface PivotCache {
        getIndex(): int
        refresh(): void
        getRecordCount(): int
        isConnected(): boolean
        getMemoryUsed(): int
        setSourceData(sourceData: any): void
        getSourceType(): int
        isOLAP(): boolean
        isBackgroundQuery(): boolean
        setBackgroundQuery(backgroundQuery: boolean): void
        getLocalConnection(): any
        setLocalConnection(localConnection: any): void
        isMaintainConnection(): boolean
        setMaintainConnection(maintainConnection: boolean): void
        getMissingItemsLimit(): int
        setMissingItemsLimit(missingItemsLimit: int): void
        isRefreshOnFileOpen(): boolean
        setRefreshOnFileOpen(refreshOnFileOpen: boolean): void
        getSourceConnectionFile(): string
        setSourceConnectionFile(sourceConnectionFile: string): void
        getSourceDataFile(): string
        isUseLocalConnection(): boolean
        setUseLocalConnection(useLocalConnection: boolean): void
        getADOConnection(): any
        getCommandText(): any
        setCommandText(commandText: any): void
        getCommandType(): int
        setCommandType(commandType: int): void
        getConnection(): any
        setConnection(connection: any): void
        isEnableRefresh(): boolean
        setEnableRefresh(enableRefresh: boolean): void
        isOptimizeCache(): boolean
        setOptimizeCache(optimizeCache: boolean): void
        getQueryType(): int
        getRecordset(): any
        setRecordset(recordset: any): void
        getRefreshDate(): double
        getRefreshName(): string
        getRefreshPeriod(): int
        setRefreshPeriod(refreshPeriod: int): void
        getRobustConnect(): int
        setRobustConnect(robustConnect: int): void
        isSavePassword(): boolean
        setSavePassword(savePassword: boolean): void
        getSourceData(): any
        createPivotTable(tableDestination: any, tableName: any, readData: any, defaultVersion: any): excel.PivotTable
        makeConnection(): void
        resetTimer(): void
        saveAsODC(ODCFileName: string, description: any, keywords: any): void
    }
    export interface PivotCache {
        getIndex(): int
        refresh(): void
        getRecordCount(): int
        isConnected(): boolean
        getMemoryUsed(): int
        setSourceData(sourceData: any): void
        getSourceType(): int
        isOLAP(): boolean
        isBackgroundQuery(): boolean
        setBackgroundQuery(backgroundQuery: boolean): void
        getLocalConnection(): any
        setLocalConnection(localConnection: any): void
        isMaintainConnection(): boolean
        setMaintainConnection(maintainConnection: boolean): void
        getMissingItemsLimit(): int
        setMissingItemsLimit(missingItemsLimit: int): void
        isRefreshOnFileOpen(): boolean
        setRefreshOnFileOpen(refreshOnFileOpen: boolean): void
        getSourceConnectionFile(): string
        setSourceConnectionFile(sourceConnectionFile: string): void
        getSourceDataFile(): string
        isUseLocalConnection(): boolean
        setUseLocalConnection(useLocalConnection: boolean): void
        getADOConnection(): any
        getCommandText(): any
        setCommandText(commandText: any): void
        getCommandType(): int
        setCommandType(commandType: int): void
        getConnection(): any
        setConnection(connection: any): void
        isEnableRefresh(): boolean
        setEnableRefresh(enableRefresh: boolean): void
        isOptimizeCache(): boolean
        setOptimizeCache(optimizeCache: boolean): void
        getQueryType(): int
        getRecordset(): any
        setRecordset(recordset: any): void
        getRefreshDate(): double
        getRefreshName(): string
        getRefreshPeriod(): int
        setRefreshPeriod(refreshPeriod: int): void
        getRobustConnect(): int
        setRobustConnect(robustConnect: int): void
        isSavePassword(): boolean
        setSavePassword(savePassword: boolean): void
        getSourceData(): any
        createPivotTable(tableDestination: any, tableName: any, readData: any, defaultVersion: any): excel.PivotTable
        makeConnection(): void
        resetTimer(): void
        saveAsODC(ODCFileName: string, description: any, keywords: any): void
    }
    export interface PivotCaches {
        add(sourceType: int, sourceData: any): excel.PivotCache
        item(index: int): excel.PivotCache
        getCount(): int
    }
    export interface PivotCaches {
        add(sourceType: int, sourceData: any): excel.PivotCache
        item(index: int): excel.PivotCache
        getCount(): int
    }
    export interface PivotCell {
        getRange(): excel.Range
        getCustomSubtotalFunction(): int
        getColumnItems(): excel.PivotItemList
        getDataField(): excel.PivotField
        getPivotCellType(): int
        getPivotField(): excel.PivotField
        getPivotItem(): excel.PivotItem
        getPivotTable(): excel.PivotTable
        getRowItems(): excel.PivotItemList
    }
    export interface PivotCell {
        getRange(): excel.Range
        getCustomSubtotalFunction(): int
        getColumnItems(): excel.PivotItemList
        getDataField(): excel.PivotField
        getPivotCellType(): int
        getPivotField(): excel.PivotField
        getPivotItem(): excel.PivotItem
        getPivotTable(): excel.PivotTable
        getRowItems(): excel.PivotItemList
    }
    export interface PivotField {
        getName(): string
        getValue(): string
        delete(): void
        setName(name: string): void
        setValue(value: string): void
        getOrientation(): int
        setCaption(caption: string): void
        getCaption(): string
        setPosition(position: any): void
        getPosition(): any
        setOrientation(orientation: int): void
        getSourceName(): string
        getNumberFormat(): string
        setNumberFormat(numberFormat: string): void
        setCurrentPage(currentPage: any): void
        setFormula(formula: string): void
        setCalculation(calculation: int): void
        getCalculation(): int
        getMemoryUsed(): int
        getLayoutSubtotalLocation(): int
        setLayoutSubtotalLocation(layoutSubtotalLocation: int): void
        setDragToColumn(dragToColumn: boolean): void
        isDragToColumn(): boolean
        isDragToData(): boolean
        setDragToData(dragToData: boolean): void
        isDragToHide(): boolean
        setDragToHide(dragToHide: boolean): void
        isDragToPage(): boolean
        setDragToPage(dragToPage: boolean): void
        isDragToRow(): boolean
        setDragToRow(dragToRow: boolean): void
        getLayoutForm(): int
        setLayoutForm(layoutForm: int): void
        autoShow(type: int, range: int, count: int, field: string): void
        autoSort(order: int, field: string, pivotLine: any, customSubtotal: any): void
        getCurrentPageList(): any
        setCurrentPageList(currentPageList: any): void
        getCurrentPageName(): string
        setCurrentPageName(currentPageName: string): void
        isEnableItemSelection(): boolean
        setEnableItemSelection(enableItemSelection: boolean): void
        getHiddenItemsList(): any
        setHiddenItemsList(hiddenItemsList: any): void
        isLayoutBlankLine(): boolean
        setLayoutBlankLine(layoutBlankLine: boolean): void
        isLayoutPageBreak(): boolean
        setLayoutPageBreak(layoutPageBreak: boolean): void
        getPropertyParentField(): excel.PivotField
        getStandardFormula(): string
        setStandardFormula(standardFormula: string): void
        getFormula(): string
        getDataType(): int
        getAutoShowCount(): int
        getAutoShowField(): string
        getAutoShowRange(): int
        getAutoShowType(): int
        getAutoSortField(): string
        getAutoSortOrder(): int
        getBaseField(): any
        setBaseField(baseField: any): void
        getBaseItem(): any
        setBaseItem(baseItem: any): void
        getChildField(): excel.PivotField
        getChildItems(index: any): any
        getCubeField(): excel.CubeField
        getCurrentPage(): any
        isDatabaseSort(): boolean
        setDatabaseSort(databaseSort: boolean): void
        getDataRange(): excel.Range
        isDrilledDown(): boolean
        setDrilledDown(drilledDown: boolean): void
        getFunction(): int
        setFunction(function: int): void
        getGroupLevel(): any
        getHiddenItems(index: any): any
        isCalculated(): boolean
        isMemberProperty(): boolean
        getLabelRange(): excel.Range
        getParentField(): excel.PivotField
        getParentItems(index: any): any
        getPropertyOrder(): int
        setPropertyOrder(propertyOrder: int): void
        isServerBased(): boolean
        setServerBased(serverBased: boolean): void
        isShowAllItems(): boolean
        setShowAllItems(showAllItems: boolean): void
        getSubtotalName(): string
        setSubtotalName(subtotalName: string): void
        getSubtotals(index: any): any
        setSubtotals(index: any, subtotals: any): void
        getTotalLevels(): any
        getVisibleItems(index: any): any
        addPageItem(item: string, clearList: any): void
        calculatedItems(): excel.CalculatedItems
        pivotItems(index: any): any
    }
    export interface PivotField {
        getName(): string
        getValue(): string
        delete(): void
        setName(name: string): void
        setValue(value: string): void
        getOrientation(): int
        setCaption(caption: string): void
        getCaption(): string
        setPosition(position: any): void
        getPosition(): any
        setOrientation(orientation: int): void
        getSourceName(): string
        getNumberFormat(): string
        setNumberFormat(numberFormat: string): void
        setCurrentPage(currentPage: any): void
        setFormula(formula: string): void
        setCalculation(calculation: int): void
        getCalculation(): int
        getMemoryUsed(): int
        getLayoutSubtotalLocation(): int
        setLayoutSubtotalLocation(layoutSubtotalLocation: int): void
        setDragToColumn(dragToColumn: boolean): void
        isDragToColumn(): boolean
        isDragToData(): boolean
        setDragToData(dragToData: boolean): void
        isDragToHide(): boolean
        setDragToHide(dragToHide: boolean): void
        isDragToPage(): boolean
        setDragToPage(dragToPage: boolean): void
        isDragToRow(): boolean
        setDragToRow(dragToRow: boolean): void
        getLayoutForm(): int
        setLayoutForm(layoutForm: int): void
        autoShow(type: int, range: int, count: int, field: string): void
        autoSort(order: int, field: string, pivotLine: any, customSubtotal: any): void
        getCurrentPageList(): any
        setCurrentPageList(currentPageList: any): void
        getCurrentPageName(): string
        setCurrentPageName(currentPageName: string): void
        isEnableItemSelection(): boolean
        setEnableItemSelection(enableItemSelection: boolean): void
        getHiddenItemsList(): any
        setHiddenItemsList(hiddenItemsList: any): void
        isLayoutBlankLine(): boolean
        setLayoutBlankLine(layoutBlankLine: boolean): void
        isLayoutPageBreak(): boolean
        setLayoutPageBreak(layoutPageBreak: boolean): void
        getPropertyParentField(): excel.PivotField
        getStandardFormula(): string
        setStandardFormula(standardFormula: string): void
        getFormula(): string
        getDataType(): int
        getAutoShowCount(): int
        getAutoShowField(): string
        getAutoShowRange(): int
        getAutoShowType(): int
        getAutoSortField(): string
        getAutoSortOrder(): int
        getBaseField(): any
        setBaseField(baseField: any): void
        getBaseItem(): any
        setBaseItem(baseItem: any): void
        getChildField(): excel.PivotField
        getChildItems(index: any): any
        getCubeField(): excel.CubeField
        getCurrentPage(): any
        isDatabaseSort(): boolean
        setDatabaseSort(databaseSort: boolean): void
        getDataRange(): excel.Range
        isDrilledDown(): boolean
        setDrilledDown(drilledDown: boolean): void
        getFunction(): int
        setFunction(function: int): void
        getGroupLevel(): any
        getHiddenItems(index: any): any
        isCalculated(): boolean
        isMemberProperty(): boolean
        getLabelRange(): excel.Range
        getParentField(): excel.PivotField
        getParentItems(index: any): any
        getPropertyOrder(): int
        setPropertyOrder(propertyOrder: int): void
        isServerBased(): boolean
        setServerBased(serverBased: boolean): void
        isShowAllItems(): boolean
        setShowAllItems(showAllItems: boolean): void
        getSubtotalName(): string
        setSubtotalName(subtotalName: string): void
        getSubtotals(index: any): any
        setSubtotals(index: any, subtotals: any): void
        getTotalLevels(): any
        getVisibleItems(index: any): any
        addPageItem(item: string, clearList: any): void
        calculatedItems(): excel.CalculatedItems
        pivotItems(index: any): any
    }
    export interface PivotFields {
        item(index: any): any
        getCount(): int
    }
    export interface PivotFields {
        item(index: any): any
        getCount(): int
    }
    export interface PivotFormula {
        getValue(): string
        delete(): void
        setValue(value: string): void
        getIndex(): int
        setIndex(index: int): void
        setFormula(formula: string): void
        getStandardFormula(): string
        setStandardFormula(standardFormula: string): void
        getFormula(): string
    }
    export interface PivotFormula {
        getValue(): string
        delete(): void
        setValue(value: string): void
        getIndex(): int
        setIndex(index: int): void
        setFormula(formula: string): void
        getStandardFormula(): string
        setStandardFormula(standardFormula: string): void
        getFormula(): string
    }
    export interface PivotFormulas {
        add(formula: string, useStandardFormula: any): excel.PivotFormula
        item(index: any): excel.PivotFormula
        getCount(): int
    }
    export interface PivotFormulas {
        add(formula: string, useStandardFormula: any): excel.PivotFormula
        item(index: any): excel.PivotFormula
        getCount(): int
    }
    export interface PivotItem {
        getName(): string
        getValue(): string
        delete(): void
        setName(name: string): void
        setValue(value: string): void
        setCaption(caption: string): void
        getCaption(): string
        isVisible(): boolean
        setVisible(visible: boolean): void
        setPosition(position: int): void
        getPosition(): int
        getSourceName(): string
        getRecordCount(): int
        setFormula(formula: string): void
        getStandardFormula(): string
        setStandardFormula(standardFormula: string): void
        isParentShowDetail(): boolean
        getSourceNameStandard(): string
        getFormula(): string
        getChildItems(index: any): any
        getDataRange(): excel.Range
        isDrilledDown(): boolean
        setDrilledDown(drilledDown: boolean): void
        isCalculated(): boolean
        getLabelRange(): excel.Range
        getParentItem(): excel.PivotItem
        isShowDetail(): boolean
        setShowDetail(showDetail: boolean): void
    }
    export interface PivotItem {
        getName(): string
        getValue(): string
        delete(): void
        setName(name: string): void
        setValue(value: string): void
        setCaption(caption: string): void
        getCaption(): string
        isVisible(): boolean
        setVisible(visible: boolean): void
        setPosition(position: int): void
        getPosition(): int
        getSourceName(): string
        getRecordCount(): int
        setFormula(formula: string): void
        getStandardFormula(): string
        setStandardFormula(standardFormula: string): void
        isParentShowDetail(): boolean
        getSourceNameStandard(): string
        getFormula(): string
        getChildItems(index: any): any
        getDataRange(): excel.Range
        isDrilledDown(): boolean
        setDrilledDown(drilledDown: boolean): void
        isCalculated(): boolean
        getLabelRange(): excel.Range
        getParentItem(): excel.PivotItem
        isShowDetail(): boolean
        setShowDetail(showDetail: boolean): void
    }
    export interface PivotItemList {
        item(index: any): excel.PivotItem
        getCount(): int
    }
    export interface PivotItemList {
        item(index: any): excel.PivotItem
        getCount(): int
    }
    export interface PivotItems {
        add(name: string): void
        item(index: any): excel.PivotItem
        getCount(): int
    }
    export interface PivotItems {
        add(name: string): void
        item(index: any): excel.PivotItem
        getCount(): int
    }
    export interface PivotLayout {
        getPivotTable(): excel.PivotTable
    }
    export interface PivotLayout {
        getPivotTable(): excel.PivotTable
    }
    export interface PivotTable {
        update(): void
        getName(): string
        format(format: int): void
        getValue(): string
        setName(name: string): void
        setValue(value: string): void
        getData(name: string): double
        getVersion(): int
        getTag(): string
        setTag(tag: string): void
        getDataFields(index: any): any
        setSelectionMode(selectionMode: int): void
        setEnableDataValueEditing(enableDataValueEditing: boolean): void
        getPivotSelectionStandard(): string
        setPivotSelectionStandard(pivotSelectionStandard: string): void
        isRepeatItemsOnEachPrintedPage(): boolean
        setRepeatItemsOnEachPrintedPage(repeatItemsOnEachPrintedPage: boolean): void
        isShowCellBackgroundFromOLAP(): boolean
        setShowCellBackgroundFromOLAP(showCellBackgroundFromOLAP: boolean): void
        isShowPageMultipleItemLabel(): boolean
        setShowPageMultipleItemLabel(showPageMultipleItemLabel: boolean): void
        isSubtotalHiddenPageItems(): boolean
        setSubtotalHiddenPageItems(subtotalHiddenPageItems: boolean): void
        setSourceData(sourceData: any): void
        getDataBodyRange(): excel.Range
        getMDX(): string
        getCalculatedMembers(): excel.CalculatedMembers
        getDataLabelRange(): excel.Range
        getDataPivotField(): excel.PivotField
        isDisplayEmptyColumn(): boolean
        setDisplayEmptyColumn(displayEmptyColumn: boolean): void
        isDisplayEmptyRow(): boolean
        setDisplayEmptyRow(displayEmptyRow: boolean): void
        isDisplayErrorString(): boolean
        setDisplayErrorString(displayErrorString: boolean): void
        isDisplayImmediateItems(): boolean
        setDisplayImmediateItems(displayImmediateItems: boolean): void
        isDisplayNullString(): boolean
        setDisplayNullString(displayNullString: boolean): void
        isEnableDataValueEditing(): boolean
        isEnableDrilldown(): boolean
        setEnableDrilldown(enableDrilldown: boolean): void
        isEnableFieldDialog(): boolean
        setEnableFieldDialog(enableFieldDialog: boolean): void
        isEnableFieldList(): boolean
        setEnableFieldList(enableFieldList: boolean): void
        getGrandTotalName(): string
        setGrandTotalName(grandTotalName: string): void
        getPageFieldOrder(): int
        setPageFieldOrder(pageFieldOrder: int): void
        getPageFieldStyle(): string
        setPageFieldStyle(pageFieldStyle: string): void
        getPageFieldWrapCount(): int
        setPageFieldWrapCount(pageFieldWrapCount: int): void
        getPageRangeCells(): excel.Range
        getPivotSelection(): string
        setPivotSelection(pivotSelection: string): void
        isPreserveFormatting(): boolean
        setPreserveFormatting(preserveFormatting: boolean): void
        isTotalsAnnotation(): boolean
        setTotalsAnnotation(totalsAnnotation: boolean): void
        isViewCalculatedMembers(): boolean
        setViewCalculatedMembers(viewCalculatedMembers: boolean): void
        getErrorString(): string
        getRefreshDate(): double
        getRefreshName(): string
        getSourceData(): any
        getCacheIndex(): int
        setCacheIndex(cacheIndex: int): void
        getColumnFields(index: any): any
        isColumnGrand(): boolean
        setColumnGrand(columnGrand: boolean): void
        getColumnRange(): excel.Range
        getCubeFields(): excel.CubeFields
        isEnableWizard(): boolean
        setEnableWizard(enableWizard: boolean): void
        setErrorString(errorString: string): void
        isHasAutoFormat(): boolean
        setHasAutoFormat(hasAutoFormat: boolean): void
        getHiddenFields(index: any): any
        getInnerDetail(): string
        setInnerDetail(innerDetail: string): void
        isManualUpdate(): boolean
        setManualUpdate(manualUpdate: boolean): void
        isMergeLabels(): boolean
        setMergeLabels(mergeLabels: boolean): void
        getNullString(): string
        setNullString(nullString: string): void
        getPageFields(index: any): any
        getPageRange(): excel.Range
        getPivotFormulas(): excel.PivotFormulas
        isPrintTitles(): boolean
        setPrintTitles(printTitles: boolean): void
        getRowFields(index: any): any
        isRowGrand(): boolean
        setRowGrand(rowGrand: boolean): void
        getRowRange(): excel.Range
        isSaveData(): boolean
        setSaveData(saveData: boolean): void
        getSelectionMode(): int
        isSmallGrid(): boolean
        setSmallGrid(smallGrid: boolean): void
        getTableRange1(): excel.Range
        getTableRange2(): excel.Range
        getTableStyle(): string
        setTableStyle(tableStyle: string): void
        getVacatedStyle(): string
        setVacatedStyle(vacatedStyle: string): void
        getVisibleFields(index: any): any
        isVisualTotals(): boolean
        setVisualTotals(visualTotals: boolean): void
        addDataField(field: any, caption: any,function: any):excel.PivotField
    addFields(rowFields: any, columnFields: any, pageFields: any, addToTable: any): any
    calculatedFields(): excel.CalculatedFields
    createCubeFile(file: string, measures: any, levels: any, members: any, properties: any): string
    getPivotData(dataField: any, field1: any, item1: any, field2: any, item2: any, field3: any, item3: any, field4: any, item4: any, field5: any, item5: any, field6: any, item6: any, field7: any, item7: any, field8: any, item8: any, field9: any, item9: any, field10: any, item10: any, field11: any, item11: any, field12: any, item12: any, field13: any, item13: any, field14: any, item14: any): excel.Range
    listFormulas(): void
        pivotCache(): excel.PivotCache
    pivotFields(index: any): any
    pivotSelect(name: string, mode: int, useStandardName: any): void
        pivotTableWizard(sourceType: any, sourceData: any, tableDestination: any, tableName: any, rowGrand: any, columnGrand: any, saveData: any, hasAutoFormat: any, autoPage: any, reserved: any, backgroundQuery: any, optimizeCache: any, pageFieldOrder: any, pageFieldWrapCount: any, readData: any, connection: any): void
            refreshTable(): boolean
    showPages(pageField: any): any

    export interface PivotTable {
        update(): void
        getName(): string
        format(format: int): void
        getValue(): string
        setName(name: string): void
        setValue(value: string): void
        getData(name: string): double
        getVersion(): int
        getTag(): string
        setTag(tag: string): void
        getDataFields(index: any): any
        setSelectionMode(selectionMode: int): void
        setEnableDataValueEditing(enableDataValueEditing: boolean): void
        getPivotSelectionStandard(): string
        setPivotSelectionStandard(pivotSelectionStandard: string): void
        isRepeatItemsOnEachPrintedPage(): boolean
        setRepeatItemsOnEachPrintedPage(repeatItemsOnEachPrintedPage: boolean): void
        isShowCellBackgroundFromOLAP(): boolean
        setShowCellBackgroundFromOLAP(showCellBackgroundFromOLAP: boolean): void
        isShowPageMultipleItemLabel(): boolean
        setShowPageMultipleItemLabel(showPageMultipleItemLabel: boolean): void
        isSubtotalHiddenPageItems(): boolean
        setSubtotalHiddenPageItems(subtotalHiddenPageItems: boolean): void
        setSourceData(sourceData: any): void
        getDataBodyRange(): excel.Range
        getMDX(): string
        getCalculatedMembers(): excel.CalculatedMembers
        getDataLabelRange(): excel.Range
        getDataPivotField(): excel.PivotField
        isDisplayEmptyColumn(): boolean
        setDisplayEmptyColumn(displayEmptyColumn: boolean): void
        isDisplayEmptyRow(): boolean
        setDisplayEmptyRow(displayEmptyRow: boolean): void
        isDisplayErrorString(): boolean
        setDisplayErrorString(displayErrorString: boolean): void
        isDisplayImmediateItems(): boolean
        setDisplayImmediateItems(displayImmediateItems: boolean): void
        isDisplayNullString(): boolean
        setDisplayNullString(displayNullString: boolean): void
        isEnableDataValueEditing(): boolean
        isEnableDrilldown(): boolean
        setEnableDrilldown(enableDrilldown: boolean): void
        isEnableFieldDialog(): boolean
        setEnableFieldDialog(enableFieldDialog: boolean): void
        isEnableFieldList(): boolean
        setEnableFieldList(enableFieldList: boolean): void
        getGrandTotalName(): string
        setGrandTotalName(grandTotalName: string): void
        getPageFieldOrder(): int
        setPageFieldOrder(pageFieldOrder: int): void
        getPageFieldStyle(): string
        setPageFieldStyle(pageFieldStyle: string): void
        getPageFieldWrapCount(): int
        setPageFieldWrapCount(pageFieldWrapCount: int): void
        getPageRangeCells(): excel.Range
        getPivotSelection(): string
        setPivotSelection(pivotSelection: string): void
        isPreserveFormatting(): boolean
        setPreserveFormatting(preserveFormatting: boolean): void
        isTotalsAnnotation(): boolean
        setTotalsAnnotation(totalsAnnotation: boolean): void
        isViewCalculatedMembers(): boolean
        setViewCalculatedMembers(viewCalculatedMembers: boolean): void
        getErrorString(): string
        getRefreshDate(): double
        getRefreshName(): string
        getSourceData(): any
        getCacheIndex(): int
        setCacheIndex(cacheIndex: int): void
        getColumnFields(index: any): any
        isColumnGrand(): boolean
        setColumnGrand(columnGrand: boolean): void
        getColumnRange(): excel.Range
        getCubeFields(): excel.CubeFields
        isEnableWizard(): boolean
        setEnableWizard(enableWizard: boolean): void
        setErrorString(errorString: string): void
        isHasAutoFormat(): boolean
        setHasAutoFormat(hasAutoFormat: boolean): void
        getHiddenFields(index: any): any
        getInnerDetail(): string
        setInnerDetail(innerDetail: string): void
        isManualUpdate(): boolean
        setManualUpdate(manualUpdate: boolean): void
        isMergeLabels(): boolean
        setMergeLabels(mergeLabels: boolean): void
        getNullString(): string
        setNullString(nullString: string): void
        getPageFields(index: any): any
        getPageRange(): excel.Range
        getPivotFormulas(): excel.PivotFormulas
        isPrintTitles(): boolean
        setPrintTitles(printTitles: boolean): void
        getRowFields(index: any): any
        isRowGrand(): boolean
        setRowGrand(rowGrand: boolean): void
        getRowRange(): excel.Range
        isSaveData(): boolean
        setSaveData(saveData: boolean): void
        getSelectionMode(): int
        isSmallGrid(): boolean
        setSmallGrid(smallGrid: boolean): void
        getTableRange1(): excel.Range
        getTableRange2(): excel.Range
        getTableStyle(): string
        setTableStyle(tableStyle: string): void
        getVacatedStyle(): string
        setVacatedStyle(vacatedStyle: string): void
        getVisibleFields(index: any): any
        isVisualTotals(): boolean
        setVisualTotals(visualTotals: boolean): void
        addDataField(field: any, caption: any,function: any):excel.PivotField
    addFields(rowFields: any, columnFields: any, pageFields: any, addToTable: any): any
    calculatedFields(): excel.CalculatedFields
    createCubeFile(file: string, measures: any, levels: any, members: any, properties: any): string
    getPivotData(dataField: any, field1: any, item1: any, field2: any, item2: any, field3: any, item3: any, field4: any, item4: any, field5: any, item5: any, field6: any, item6: any, field7: any, item7: any, field8: any, item8: any, field9: any, item9: any, field10: any, item10: any, field11: any, item11: any, field12: any, item12: any, field13: any, item13: any, field14: any, item14: any): excel.Range
    listFormulas(): void
        pivotCache(): excel.PivotCache
    pivotFields(index: any): any
    pivotSelect(name: string, mode: int, useStandardName: any): void
        pivotTableWizard(sourceType: any, sourceData: any, tableDestination: any, tableName: any, rowGrand: any, columnGrand: any, saveData: any, hasAutoFormat: any, autoPage: any, reserved: any, backgroundQuery: any, optimizeCache: any, pageFieldOrder: any, pageFieldWrapCount: any, readData: any, connection: any): void
            refreshTable(): boolean
    showPages(pageField: any): any

    export interface PivotTables {
        add(pivotCache: excel.PivotCache, tableDestination: excel.Range, tableName: string, readData: boolean, defaultVersion: int): excel.PivotTable
        item(index: any): excel.PivotTable
        getCount(): int
    }
    export interface PivotTables {
        add(pivotCache: excel.PivotCache, tableDestination: excel.Range, tableName: string, readData: boolean, defaultVersion: int): excel.PivotTable
        item(index: any): excel.PivotTable
        getCount(): int
    }
    export interface PlotArea {
        getName(): string
        getBorder(): excel.Border
        getWidth(): double
        getLeft(): double
        setLeft(left: double): void
        setTop(top: double): void
        getTop(): double
        setWidth(width: double): void
        getHeight(): double
        setHeight(height: double): void
        select(): any
        getFill(): excel.ChartFillFormat
        clearFormats(): any
        getInterior(): excel.Interior
        getInsideHeight(): double
        getInsideLeft(): double
        getInsideTop(): double
        getInsideWidth(): double
    }
    export interface PlotArea {
        getName(): string
        getBorder(): excel.Border
        getWidth(): double
        getLeft(): double
        setLeft(left: double): void
        setTop(top: double): void
        getTop(): double
        setWidth(width: double): void
        getHeight(): double
        setHeight(height: double): void
        select(): any
        getFill(): excel.ChartFillFormat
        clearFormats(): any
        getInterior(): excel.Interior
        getInsideHeight(): double
        getInsideLeft(): double
        getInsideTop(): double
        getInsideWidth(): double
    }
    export interface Point {
        delete(): any
        copy(): any
        getBorder(): excel.Border
        paste(): any
        select(): any
        isShadow(): boolean
        setShadow(shadow: boolean): void
        getFill(): excel.ChartFillFormat
        clearFormats(): any
        getInterior(): excel.Interior
        getPictureType(): int
        setPictureType(pictureType: int): void
        setMarkerSize(markerSize: int): void
        getMarkerSize(): int
        setMarkerStyle(markerStyle: int): void
        getMarkerStyle(): int
        setPictureUnit(pictureUnit: int): void
        getPictureUnit(): int
        setApplyPictToEnd(applyPictToEnd: boolean): void
        isApplyPictToFront(): boolean
        setApplyPictToFront(applyPictToFront: boolean): void
        isApplyPictToSides(): boolean
        setApplyPictToSides(applyPictToSides: boolean): void
        applyDataLabels(type: any, legendKey: any, autoText: any, hasLeaderLines: any, showSeriesName: any, showCategoryName: any, showValue: any, showPercentage: any, showBubbleSize: any, separator: any): any
        setInvertIfNegative(invertIfNegative: boolean): void
        isInvertIfNegative(): boolean
        setMarkerBackgroundColor(markerBackgroundColor: int): void
        getMarkerBackgroundColor(): int
        setMarkerForegroundColor(markerForegroundColor: int): void
        getMarkerForegroundColor(): int
        setMarkerBackgroundColorIndex(markerBackgroundColorIndex: int): void
        getMarkerBackgroundColorIndex(): int
        setMarkerForegroundColorIndex(markerForegroundColorIndex: int): void
        getMarkerForegroundColorIndex(): int
        isApplyPictToEnd(): boolean
        getDataLabel(): excel.DataLabel
        getExplosion(): int
        setExplosion(explosion: int): void
        isHasDataLabel(): boolean
        setHasDataLabel(hasDataLabel: boolean): void
        isSecondaryPlot(): boolean
        setSecondaryPlot(secondaryPlot: boolean): void
    }
    export interface Point {
        delete(): any
        copy(): any
        getBorder(): excel.Border
        paste(): any
        select(): any
        isShadow(): boolean
        setShadow(shadow: boolean): void
        getFill(): excel.ChartFillFormat
        clearFormats(): any
        getInterior(): excel.Interior
        getPictureType(): int
        setPictureType(pictureType: int): void
        setMarkerSize(markerSize: int): void
        getMarkerSize(): int
        setMarkerStyle(markerStyle: int): void
        getMarkerStyle(): int
        setPictureUnit(pictureUnit: int): void
        getPictureUnit(): int
        setApplyPictToEnd(applyPictToEnd: boolean): void
        isApplyPictToFront(): boolean
        setApplyPictToFront(applyPictToFront: boolean): void
        isApplyPictToSides(): boolean
        setApplyPictToSides(applyPictToSides: boolean): void
        applyDataLabels(type: any, legendKey: any, autoText: any, hasLeaderLines: any, showSeriesName: any, showCategoryName: any, showValue: any, showPercentage: any, showBubbleSize: any, separator: any): any
        setInvertIfNegative(invertIfNegative: boolean): void
        isInvertIfNegative(): boolean
        setMarkerBackgroundColor(markerBackgroundColor: int): void
        getMarkerBackgroundColor(): int
        setMarkerForegroundColor(markerForegroundColor: int): void
        getMarkerForegroundColor(): int
        setMarkerBackgroundColorIndex(markerBackgroundColorIndex: int): void
        getMarkerBackgroundColorIndex(): int
        setMarkerForegroundColorIndex(markerForegroundColorIndex: int): void
        getMarkerForegroundColorIndex(): int
        isApplyPictToEnd(): boolean
        getDataLabel(): excel.DataLabel
        getExplosion(): int
        setExplosion(explosion: int): void
        isHasDataLabel(): boolean
        setHasDataLabel(hasDataLabel: boolean): void
        isSecondaryPlot(): boolean
        setSecondaryPlot(secondaryPlot: boolean): void
    }
    export interface Points {
        item(index: int): excel.Point
        getCount(): int
    }
    export interface Points {
        item(index: int): excel.Point
        getCount(): int
    }
    export interface Protection {
        isAllowInsertingHyperlinks(): boolean
        resetMProtectProperties(protect: any): void
        isAllowDeletingColumns(): boolean
        isAllowDeletingRows(): boolean
        getAllowEditRanges(): excel.AllowEditRanges
        isAllowFormattingCells(): boolean
        isAllowFormattingColumns(): boolean
        isAllowFormattingRows(): boolean
        isAllowInsertingColumns(): boolean
        isAllowInsertingRows(): boolean
        isAllowUsingPivotTables(): boolean
        isAllowFiltering(): boolean
        isAllowSorting(): boolean
    }
    export interface Protection {
        isAllowInsertingHyperlinks(): boolean
        resetMProtectProperties(protect: any): void
        isAllowDeletingColumns(): boolean
        isAllowDeletingRows(): boolean
        getAllowEditRanges(): excel.AllowEditRanges
        isAllowFormattingCells(): boolean
        isAllowFormattingColumns(): boolean
        isAllowFormattingRows(): boolean
        isAllowInsertingColumns(): boolean
        isAllowInsertingRows(): boolean
        isAllowUsingPivotTables(): boolean
        isAllowFiltering(): boolean
        isAllowSorting(): boolean
    }
    export interface PublishObject {
        delete(): void
        setTitle(title: string): void
        getTitle(): string
        getSource(): string
        getSheet(): string
        getSourceType(): int
        getDivID(): string
        setDivID(divID: string): void
        publish(create: boolean): void
        isAutoRepublish(): boolean
        setAutoRepublish(autoRepublish: boolean): void
        getFilename(): string
        setFilename(filename: string): void
        getHtmlType(): int
        setHtmlType(htmlType: int): void
    }
    export interface PublishObject {
        delete(): void
        setTitle(title: string): void
        getTitle(): string
        getSource(): string
        getSheet(): string
        getSourceType(): int
        getDivID(): string
        setDivID(divID: string): void
        publish(create: boolean): void
        isAutoRepublish(): boolean
        setAutoRepublish(autoRepublish: boolean): void
        getFilename(): string
        setFilename(filename: string): void
        getHtmlType(): int
        setHtmlType(htmlType: int): void
    }
    export interface PublishObjects {
        add(sourceType: int, filename: string, sheet: any, source: any, htmlType: any, divID: any, title: any): excel.PublishObject
        delete(): void
        item(index: any): excel.PublishObject
        getCount(): int
        publish(): void
    }
    export interface PublishObjects {
        add(sourceType: int, filename: string, sheet: any, source: any, htmlType: any, divID: any, title: any): excel.PublishObject
        delete(): void
        item(index: any): excel.PublishObject
        getCount(): int
        publish(): void
    }
    export interface QueryTable {
        getName(): string
        delete(): void
        setName(name: string): void
        getParameters(): excel.Parameters
        refresh(backgroundQuery: any): boolean
        getDestination(): excel.Range
        getTextFileColumnDataTypes(): any
        setTextFileColumnDataTypes(dataType: any): void
        setTextFileCommaDelimiter(textFileCommaDelimiter: boolean): void
        isTextFileConsecutiveDelimiter(): boolean
        setTextFileConsecutiveDelimiter(textFileConsecutiveDelimiter: boolean): void
        getTextFileDecimalSeparator(): string
        setTextFileDecimalSeparator(textFileDecimalSeparator: string): void
        getTextFileFixedColumnWidths(): any
        setTextFileFixedColumnWidths(chars: any): void
        getTextFileOtherDelimiter(): string
        setTextFileOtherDelimiter(otherDelimiter: string): void
        isTextFilePromptOnRefresh(): boolean
        setTextFilePromptOnRefresh(textFilePromptOnRefresh: boolean): void
        isTextFileSemicolonDelimiter(): boolean
        setTextFileSemicolonDelimiter(textFileSemicolonDelimiter: boolean): void
        setTextFileSpaceDelimiter(textFileSpaceDelimiter: boolean): void
        getTextFileThousandsSeparator(): string
        setTextFileThousandsSeparator(textFileThousandsSeparator: string): void
        isTextFileTrailingMinusNumbers(): boolean
        setTextFileTrailingMinusNumbers(textFileTrailingMinusNumbers: boolean): void
        isWebConsecutiveDelimitersAsOne(): boolean
        setWebConsecutiveDelimitersAsOne(webConsecutiveDelimitersAsOne: boolean): void
        isWebDisableDateRecognition(): boolean
        setWebDisableDateRecognition(webDisableDateRecognition: boolean): void
        setWebDisableRedirections(webDisableRedirections: boolean): void
        isWebPreFormattedTextToColumns(): boolean
        setWebPreFormattedTextToColumns(webPreFormattedTextToColumns: boolean): void
        isWebSingleBlockTextImport(): boolean
        setWebSingleBlockTextImport(webSingleBlockTextImport: boolean): void
        removeYzoQueryTableListener(listener: excel.event.QueryTableListener): void
        isBackgroundQuery(): boolean
        setBackgroundQuery(backgroundQuery: boolean): void
        isMaintainConnection(): boolean
        setMaintainConnection(maintainConnection: boolean): void
        isRefreshOnFileOpen(): boolean
        setRefreshOnFileOpen(refreshOnFileOpen: boolean): void
        getSourceConnectionFile(): string
        setSourceConnectionFile(sourceConnectionFile: string): void
        getSourceDataFile(): string
        isPreserveFormatting(): boolean
        setPreserveFormatting(preserveFormatting: boolean): void
        isAdjustColumnWidth(): boolean
        setAdjustColumnWidth(adjustColumnWidth: boolean): void
        isFetchedRowOverflow(): boolean
        isFillAdjacentFormulas(): boolean
        setFillAdjacentFormulas(fillAdjacentFormulas: boolean): void
        isPreserveColumnInfo(): boolean
        setPreserveColumnInfo(preserveColumnInfo: boolean): void
        setSourceDataFile(sourceDataFile: string): void
        isTextFileCommaDelimiter(): boolean
        setTextFileParseType(type: int): void
        getTextFileParseType(): int
        getTextFilePlatform(): int
        setTextFilePlatform(textFilePlatform: int): void
        isTextFileSpaceDelimiter(): boolean
        getTextFileStartRow(): int
        setTextFileStartRow(startRow: int): void
        isTextFileTabDelimiter(): boolean
        setTextFileTabDelimiter(textFileTabDelimiter: boolean): void
        getTextFileTextQualifier(): int
        setTextFileTextQualifier(type: int): void
        getTextFileVisualLayout(): int
        setTextFileVisualLayout(textFileVisualLayout: int): void
        isWebDisableRedirections(): boolean
        getWebSelectionType(): int
        setWebSelectionType(webSelectionType: int): void
        addYzoQueryTableListener(listener: excel.event.QueryTableListener): void
        getCommandText(): any
        setCommandText(commandText: any): void
        getCommandType(): int
        setCommandType(commandType: int): void
        getConnection(): any
        setConnection(connection: any): void
        isEnableRefresh(): boolean
        setEnableRefresh(enableRefresh: boolean): void
        getQueryType(): int
        getRecordset(): any
        setRecordset(recordset: any): void
        getRefreshPeriod(): int
        setRefreshPeriod(refreshPeriod: int): void
        getRobustConnect(): int
        setRobustConnect(robustConnect: int): void
        isSavePassword(): boolean
        setSavePassword(savePassword: boolean): void
        resetTimer(): void
        saveAsODC(ODCFileName: string, description: any, keywords: any): void
        isSaveData(): boolean
        setSaveData(saveData: boolean): void
        getEditWebPage(): any
        setEditWebPage(editWebPage: any): void
        isEnableEditing(): boolean
        setEnableEditing(enableEditing: boolean): void
        isFieldNames(): boolean
        setFieldNames(fieldNames: boolean): void
        getListObject(): excel.ListObject
        getPostText(): string
        setPostText(postText: string): void
        isRefreshing(): boolean
        getRefreshStyle(): int
        setRefreshStyle(style: int): void
        getResultRange(): excel.Range
        isRowNumbers(): boolean
        setRowNumbers(rowNumbers: boolean): void
        getWebFormatting(): int
        setWebFormatting(webFormatting: int): void
        getWebTables(): string
        setWebTables(webTables: string): void
        cancelRefresh(): void
    }
    export interface QueryTable {
        getName(): string
        delete(): void
        setName(name: string): void
        getParameters(): excel.Parameters
        refresh(backgroundQuery: any): boolean
        getDestination(): excel.Range
        getTextFileColumnDataTypes(): any
        setTextFileColumnDataTypes(dataType: any): void
        setTextFileCommaDelimiter(textFileCommaDelimiter: boolean): void
        isTextFileConsecutiveDelimiter(): boolean
        setTextFileConsecutiveDelimiter(textFileConsecutiveDelimiter: boolean): void
        getTextFileDecimalSeparator(): string
        setTextFileDecimalSeparator(textFileDecimalSeparator: string): void
        getTextFileFixedColumnWidths(): any
        setTextFileFixedColumnWidths(chars: any): void
        getTextFileOtherDelimiter(): string
        setTextFileOtherDelimiter(otherDelimiter: string): void
        isTextFilePromptOnRefresh(): boolean
        setTextFilePromptOnRefresh(textFilePromptOnRefresh: boolean): void
        isTextFileSemicolonDelimiter(): boolean
        setTextFileSemicolonDelimiter(textFileSemicolonDelimiter: boolean): void
        setTextFileSpaceDelimiter(textFileSpaceDelimiter: boolean): void
        getTextFileThousandsSeparator(): string
        setTextFileThousandsSeparator(textFileThousandsSeparator: string): void
        isTextFileTrailingMinusNumbers(): boolean
        setTextFileTrailingMinusNumbers(textFileTrailingMinusNumbers: boolean): void
        isWebConsecutiveDelimitersAsOne(): boolean
        setWebConsecutiveDelimitersAsOne(webConsecutiveDelimitersAsOne: boolean): void
        isWebDisableDateRecognition(): boolean
        setWebDisableDateRecognition(webDisableDateRecognition: boolean): void
        setWebDisableRedirections(webDisableRedirections: boolean): void
        isWebPreFormattedTextToColumns(): boolean
        setWebPreFormattedTextToColumns(webPreFormattedTextToColumns: boolean): void
        isWebSingleBlockTextImport(): boolean
        setWebSingleBlockTextImport(webSingleBlockTextImport: boolean): void
        removeYzoQueryTableListener(listener: excel.event.QueryTableListener): void
        isBackgroundQuery(): boolean
        setBackgroundQuery(backgroundQuery: boolean): void
        isMaintainConnection(): boolean
        setMaintainConnection(maintainConnection: boolean): void
        isRefreshOnFileOpen(): boolean
        setRefreshOnFileOpen(refreshOnFileOpen: boolean): void
        getSourceConnectionFile(): string
        setSourceConnectionFile(sourceConnectionFile: string): void
        getSourceDataFile(): string
        isPreserveFormatting(): boolean
        setPreserveFormatting(preserveFormatting: boolean): void
        isAdjustColumnWidth(): boolean
        setAdjustColumnWidth(adjustColumnWidth: boolean): void
        isFetchedRowOverflow(): boolean
        isFillAdjacentFormulas(): boolean
        setFillAdjacentFormulas(fillAdjacentFormulas: boolean): void
        isPreserveColumnInfo(): boolean
        setPreserveColumnInfo(preserveColumnInfo: boolean): void
        setSourceDataFile(sourceDataFile: string): void
        isTextFileCommaDelimiter(): boolean
        setTextFileParseType(type: int): void
        getTextFileParseType(): int
        getTextFilePlatform(): int
        setTextFilePlatform(textFilePlatform: int): void
        isTextFileSpaceDelimiter(): boolean
        getTextFileStartRow(): int
        setTextFileStartRow(startRow: int): void
        isTextFileTabDelimiter(): boolean
        setTextFileTabDelimiter(textFileTabDelimiter: boolean): void
        getTextFileTextQualifier(): int
        setTextFileTextQualifier(type: int): void
        getTextFileVisualLayout(): int
        setTextFileVisualLayout(textFileVisualLayout: int): void
        isWebDisableRedirections(): boolean
        getWebSelectionType(): int
        setWebSelectionType(webSelectionType: int): void
        addYzoQueryTableListener(listener: excel.event.QueryTableListener): void
        getCommandText(): any
        setCommandText(commandText: any): void
        getCommandType(): int
        setCommandType(commandType: int): void
        getConnection(): any
        setConnection(connection: any): void
        isEnableRefresh(): boolean
        setEnableRefresh(enableRefresh: boolean): void
        getQueryType(): int
        getRecordset(): any
        setRecordset(recordset: any): void
        getRefreshPeriod(): int
        setRefreshPeriod(refreshPeriod: int): void
        getRobustConnect(): int
        setRobustConnect(robustConnect: int): void
        isSavePassword(): boolean
        setSavePassword(savePassword: boolean): void
        resetTimer(): void
        saveAsODC(ODCFileName: string, description: any, keywords: any): void
        isSaveData(): boolean
        setSaveData(saveData: boolean): void
        getEditWebPage(): any
        setEditWebPage(editWebPage: any): void
        isEnableEditing(): boolean
        setEnableEditing(enableEditing: boolean): void
        isFieldNames(): boolean
        setFieldNames(fieldNames: boolean): void
        getListObject(): excel.ListObject
        getPostText(): string
        setPostText(postText: string): void
        isRefreshing(): boolean
        getRefreshStyle(): int
        setRefreshStyle(style: int): void
        getResultRange(): excel.Range
        isRowNumbers(): boolean
        setRowNumbers(rowNumbers: boolean): void
        getWebFormatting(): int
        setWebFormatting(webFormatting: int): void
        getWebTables(): string
        setWebTables(webTables: string): void
        cancelRefresh(): void
    }
    export interface QueryTables {
        add(connection: any, destination: excel.Range, sql: any): excel.QueryTable
        item(index: any): excel.QueryTable
        getCount(): int
    }
    export interface QueryTables {
        add(connection: any, destination: excel.Range, sql: any): excel.QueryTable
        item(index: any): excel.QueryTable
        getCount(): int
    }
    export interface Range {
        group(start: any, end: any, by: any, periods: any): any
        run(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): any
        getAddress(rowAbsolute: any, columnAbsolute: any, referenceStyle: int, external: any, relativeTo: any): string
        clear(): any
        getName(): any
        replace(what: any, replacement: any, lookAt: any, searchOrder: any, matchCase: any, matchByte: any, searchFormat: any, replaceFormat: any): boolean
        getValue(tangeValueDataType: any): any
        hasArray(): boolean
        find(what: any, after: any, lookIn: any, lookAt: any, searchOrder: any, searchDirection: int, matchCase: any, matchByte: any, serchFormat: any): excel.Range
        delete(shift: any): any
        setName(pName: any): void
        table(rowInput: any, columnInput: any): any
        merge(across: any): void
        setValue(pValue: any): void
        setValue(pValue: any, rangeValueDataType: any): void
        copy(destination: any): any
        insert(shift: any, copyOrigin: any): any
        getOffset(rowOffset: any, columnOffset: any): excel.Range
        sort(key1: any, order1: int, key2: any, type: any, order2: int, key3: any, order3: int, header: int, orderCustom: any, matchCase: any, orientation: int, sortMethod: int, dataOption1: int, dataOption2: int, dataOption3: int): any
        parse(parseLine: any, destination: any): any
        getID(): string
        item(rowIndex: any, columnIndex: any): excel.Range
        getFont(): excel.Font
        getComment(): excel.Comment
        end(direction: int): excel.Range
        getOrientation(): any
        show(): any
        getCount(): int
        getWidth(): any
        getLeft(): any
        getTop(): any
        activate(): any
        getHeight(): any
        printPreview(EnableChanges: any): any
        checkSpelling(customDictionary: any, ignoreUppercase: any, alwaysSuggest: any, spellLang: any): any
        printOut(from: any, to: any, copies: any, preview: any, activePrinter: any, printToFile: any, collate: any, prToFileName: any): any
        getText(): any
        cut(destination: any): any
        getMRange(): any
        setRowHeight(pRowHeight: any): void
        setColumnWidth(pColumnWidth: any): void
        getRowHeight(): any
        getPrevious(): excel.Range
        getEnd(direction: int): excel.Range
        getRange(cell1: any, cell2: any): excel.Range
        select(): any
        setID(pID: string): void
        getNext(): excel.Range
        getRow(): int
        getCells(row: any, column: any): excel.Range
        getCells(): excel.Range
        autoFit(): any
        getBorders(): excel.Borders
        getColumn(): int
        setVerticalAlignment(pVerticalAlignment: any): void
        getVerticalAlignment(): any
        Replace(what: any, replacement: any, lookAt: any, searchOrder: any, matchCase: any, matchByte: any, searchFormat: any, replaceFormat: any): boolean
        clearFormats(): any
        pasteSpecial(paste: int, operation: int, skipBlanks: any, transpose: any): any
        setOrientation(pOrientation: any): void
        getLocked(): any
        getColumns(): excel.Range
        getNumberFormat(): any
        setNumberFormat(pNumberFormat: any): void
        getRows(): excel.Range
        getSmartTags(): excel.SmartTags
        getStyle(): any
        setStyle(pStyle: any): void
        getHidden(): any
        setHidden(pHidden: any): void
        setLocked(pLocked: any): void
        autoFormat(format: int, number: int, font: excel.Font, alignment: int, border: excel.Border, pattern: any, width: float): void
        getCharacters(start: any, length: any): excel.Characters
        getHyperlinks(): excel.Hyperlinks
        ungroup(): any
        getOutlineLevel(): any
        setOutlineLevel(pOutlineLevel: any): void
        getReadingOrder(): int
        setReadingOrder(pReadingOrder: int): void
        subscribeTo(edition: string, format: int): any
        calculate(): any
        createPublisher(edition: any, appearance: int, containsPICT: any, containsBIFF: any, containsRTF: any, containsVALU: any): void
        getHorizontalAlignment(): any
        setHorizontalAlignment(pHorizontalAlignment: any): void
        setFormula(pFormula: any): void
        getWorksheet(): excel.Worksheet
        getInterior(): excel.Interior
        getBook(): any
        justify(): any
        autoFill(destination: excel.Range, type: int): any
        findNext(after: any): excel.Range
        subtotal(groupBy: int,function: int, totalList: any, replace: any, pageBreaks: any, summaryBelowData: int): any
        goalSeek(goal: any, changingCell: excel.Range): boolean
        getPhonetic(): excel.Phonetic
        getFormatConditions(): excel.FormatConditions
        copyPicture(appearance: int, format: int): any
        clearContents(): any
        getAddIndent(): any
        getFormulaHidden(): any
        getShrinkToFit(): any
        getWrapText(): any
        getQueryTable(): excel.QueryTable
        getSummary(): any
        getXPath(): excel.XPath
        getAreas(): excel.Areas
        dirty(): void
        fillDown(): any
        fillLeft(): any
        fillUp(): any
        noteText(text: any, start: any, length: any): string
        speak(speakDirection: any, speakFormulas: any): void
        unMerge(): void
        setFormulaArray(pFormulaArray: any): void
        getValue2(): any
        removeSubtotal(): any
        consolidate(sources: any,function: any, topRow: any, leftColumn: any, createLinks: any): any
        clearOutline(): any
        hasFormula(): any
        getFormula(): any
        setWrapText(pWrapText: any): void
        setAddIndent(pAddIndent: any): void
        setFormulaHidden(pFormulaHidden: any): void
        setIndentLevel(pIndentLevel: any): void
        getIndentLevel(): any
        setMergeCells(pMergeCells: any): void
        isMergeCells(): any
        setShrinkToFit(pShrinkToFit: any): void
        setNumberFormatLocal(pNumberFormatLocal: any): void
        getNumberFormatLocal(): any
        getDirectDependents(): excel.Range
        getDirectPrecedents(): excel.Range
        getFormulaR1C1Local(): any
        getListHeaderRows(): int
        getLocationInTable(): int
        getPrefixCharacter(): any
        getUseStandardHeight(): any
        getUseStandardWidth(): any
        setFormulaR1C1Local(pFormulaR1C1Local: any): void
        setUseStandardHeight(pUseStandardHeight: any): void
        setUseStandardWidth(pUseStandardWidth: any): void
        applyOutlineStyles(): any
        columnDifferences(comparison: any): excel.Range
        copyFromRecordset(data: any, maxRows: any, maxColumns: any): int
        getPivotField(): excel.PivotField
        getPivotItem(): excel.PivotItem
        getPivotTable(): excel.PivotTable
        setShowDetail(pShowDetail: any): void
        getListObject(): excel.ListObject
        getAddressLocal(rowAbsolute: any, columnAbsolute: any, referenceStyle: int, external: any, relativeTo: any): string
        isAllowEdit(): boolean
        getColumnWidth(): any
        getCountLarge(): any
        getCurrentArray(): excel.Range
        getCurrentRegion(): excel.Range
        getDependents(): excel.Range
        getPrecedents(): excel.Range
        getDisplayFormat(): excel.DisplayFormat
        getEntireColumn(): excel.Range
        getEntireRow(): excel.Range
        getErrors(): excel.Errors
        getFormulaArray(): any
        getFormulaLabel(): int
        getFormulaLocal(): any
        getFormulaR1C1(): any
        getMergeArea(): excel.Range
        getPageBreak(): int
        getPhonetics(): excel.Phonetics
        getPivotCell(): excel.PivotCell
        getResize(rowSize: any, columnSize: any): excel.Range
        getShowDetail(): any
        getValidation(): excel.Validation
        setValue2(pValue2: any): void
        setFormulaLabel(pFormulaLabel: int): void
        setFormulaLocal(pFormulaLocal: any): void
        setFormulaR1C1(pFormulaR1C1: any): void
        setPageBreak(pPageBreak: int): void
        addComment(text: any): excel.Comment
        advancedFilter(action: int, criteriaRange: any, copyToRange: any, unique: any): any
        applyNames(names: any, ignoreRelativeAbsolute: any, useRowColumnNames: any, omitColumn: any, omitRow: any, order: int, appendLast: any): any
        autoComplete(string: string): string
        autoFilter(field: any, criteria1: any, operator: int, criteria2: any, visibleDropDown: any): any
        paserCriteria(criteria: any): any[]
        autoOutline(): any
        borderAround(lineStyle: any, weight: int, colorIndex: int, color: any, themeColor: any): any
        clearComments(): void
        clearHyperlinks(): void
        clearNotes(): any
        createNames(top: any, left: any, bottom: any, right: any): any
        dataSeries(rowcol: any, type: int, date: int, step: any, stop: any, trend: any): any
        dialogBox(): any
        editionOptions(type: int, option: int, name: any, reference: any, appearance: int, chartSize: int, format: any): any
        fillRight(): any
        findPrevious(after: any): excel.Range
        functionWizard(): any
        insertIndent(insertAmount: int): void
        listNames(): any
        navigateArrow(towardPrecedent: any, arrowNumber: any, linkNumber: any): any
        RemoveDuplicates(columns: any, header: any): void
        rowDifferences(comparison: any): excel.Range
        setPhonetic(): void
        showDependents(remove: any): any
        showErrors(): any
        showPrecedents(remove: any): any
        sortSpecial(sortMethod: int, key1: any, order1: int, type: any, key2: any, order2: int, key3: any, order3: int, header: int, orderCustom: any, matchCase: any, orientation: int, dataOption1: int, dataOption2: int, dataOption3: int): any
        specialCells(type: int, value: any): excel.Range
        textToColumns(destination: any, dataType: int, textQualifier: int, consecutiveDelimiter: any, tab: any, semicolon: any, comma: any, space: any, other: any, otherChar: any, fieldInfo: any, decimalSeparator: any, thousandsSeparator: any, trailingMinusNumbers: any): any
    }
    export interface Range {
        group(start: any, end: any, by: any, periods: any): any
        run(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): any
        getAddress(rowAbsolute: any, columnAbsolute: any, referenceStyle: int, external: any, relativeTo: any): string
        clear(): any
        getName(): any
        replace(what: any, replacement: any, lookAt: any, searchOrder: any, matchCase: any, matchByte: any, searchFormat: any, replaceFormat: any): boolean
        getValue(tangeValueDataType: any): any
        hasArray(): boolean
        find(what: any, after: any, lookIn: any, lookAt: any, searchOrder: any, searchDirection: int, matchCase: any, matchByte: any, serchFormat: any): excel.Range
        delete(shift: any): any
        setName(pName: any): void
        table(rowInput: any, columnInput: any): any
        merge(across: any): void
        setValue(pValue: any): void
        setValue(pValue: any, rangeValueDataType: any): void
        copy(destination: any): any
        insert(shift: any, copyOrigin: any): any
        getOffset(rowOffset: any, columnOffset: any): excel.Range
        sort(key1: any, order1: int, key2: any, type: any, order2: int, key3: any, order3: int, header: int, orderCustom: any, matchCase: any, orientation: int, sortMethod: int, dataOption1: int, dataOption2: int, dataOption3: int): any
        parse(parseLine: any, destination: any): any
        getID(): string
        item(rowIndex: any, columnIndex: any): excel.Range
        getFont(): excel.Font
        getComment(): excel.Comment
        end(direction: int): excel.Range
        getOrientation(): any
        show(): any
        getCount(): int
        getWidth(): any
        getLeft(): any
        getTop(): any
        activate(): any
        getHeight(): any
        printPreview(EnableChanges: any): any
        checkSpelling(customDictionary: any, ignoreUppercase: any, alwaysSuggest: any, spellLang: any): any
        printOut(from: any, to: any, copies: any, preview: any, activePrinter: any, printToFile: any, collate: any, prToFileName: any): any
        getText(): any
        cut(destination: any): any
        getMRange(): any
        setRowHeight(pRowHeight: any): void
        setColumnWidth(pColumnWidth: any): void
        getRowHeight(): any
        getPrevious(): excel.Range
        getEnd(direction: int): excel.Range
        getRange(cell1: any, cell2: any): excel.Range
        select(): any
        setID(pID: string): void
        getNext(): excel.Range
        getRow(): int
        getCells(row: any, column: any): excel.Range
        getCells(): excel.Range
        autoFit(): any
        getBorders(): excel.Borders
        getColumn(): int
        setVerticalAlignment(pVerticalAlignment: any): void
        getVerticalAlignment(): any
        Replace(what: any, replacement: any, lookAt: any, searchOrder: any, matchCase: any, matchByte: any, searchFormat: any, replaceFormat: any): boolean
        clearFormats(): any
        pasteSpecial(paste: int, operation: int, skipBlanks: any, transpose: any): any
        setOrientation(pOrientation: any): void
        getLocked(): any
        getColumns(): excel.Range
        getNumberFormat(): any
        setNumberFormat(pNumberFormat: any): void
        getRows(): excel.Range
        getSmartTags(): excel.SmartTags
        getStyle(): any
        setStyle(pStyle: any): void
        getHidden(): any
        setHidden(pHidden: any): void
        setLocked(pLocked: any): void
        autoFormat(format: int, number: int, font: excel.Font, alignment: int, border: excel.Border, pattern: any, width: float): void
        getCharacters(start: any, length: any): excel.Characters
        getHyperlinks(): excel.Hyperlinks
        ungroup(): any
        getOutlineLevel(): any
        setOutlineLevel(pOutlineLevel: any): void
        getReadingOrder(): int
        setReadingOrder(pReadingOrder: int): void
        subscribeTo(edition: string, format: int): any
        calculate(): any
        createPublisher(edition: any, appearance: int, containsPICT: any, containsBIFF: any, containsRTF: any, containsVALU: any): void
        getHorizontalAlignment(): any
        setHorizontalAlignment(pHorizontalAlignment: any): void
        setFormula(pFormula: any): void
        getWorksheet(): excel.Worksheet
        getInterior(): excel.Interior
        getBook(): any
        justify(): any
        autoFill(destination: excel.Range, type: int): any
        findNext(after: any): excel.Range
        subtotal(groupBy: int,function: int, totalList: any, replace: any, pageBreaks: any, summaryBelowData: int): any
        goalSeek(goal: any, changingCell: excel.Range): boolean
        getPhonetic(): excel.Phonetic
        getFormatConditions(): excel.FormatConditions
        copyPicture(appearance: int, format: int): any
        clearContents(): any
        getAddIndent(): any
        getFormulaHidden(): any
        getShrinkToFit(): any
        getWrapText(): any
        getQueryTable(): excel.QueryTable
        getSummary(): any
        getXPath(): excel.XPath
        getAreas(): excel.Areas
        dirty(): void
        fillDown(): any
        fillLeft(): any
        fillUp(): any
        noteText(text: any, start: any, length: any): string
        speak(speakDirection: any, speakFormulas: any): void
        unMerge(): void
        setFormulaArray(pFormulaArray: any): void
        getValue2(): any
        removeSubtotal(): any
        consolidate(sources: any,function: any, topRow: any, leftColumn: any, createLinks: any): any
        clearOutline(): any
        hasFormula(): any
        getFormula(): any
        setWrapText(pWrapText: any): void
        setAddIndent(pAddIndent: any): void
        setFormulaHidden(pFormulaHidden: any): void
        setIndentLevel(pIndentLevel: any): void
        getIndentLevel(): any
        setMergeCells(pMergeCells: any): void
        isMergeCells(): any
        setShrinkToFit(pShrinkToFit: any): void
        setNumberFormatLocal(pNumberFormatLocal: any): void
        getNumberFormatLocal(): any
        getDirectDependents(): excel.Range
        getDirectPrecedents(): excel.Range
        getFormulaR1C1Local(): any
        getListHeaderRows(): int
        getLocationInTable(): int
        getPrefixCharacter(): any
        getUseStandardHeight(): any
        getUseStandardWidth(): any
        setFormulaR1C1Local(pFormulaR1C1Local: any): void
        setUseStandardHeight(pUseStandardHeight: any): void
        setUseStandardWidth(pUseStandardWidth: any): void
        applyOutlineStyles(): any
        columnDifferences(comparison: any): excel.Range
        copyFromRecordset(data: any, maxRows: any, maxColumns: any): int
        getPivotField(): excel.PivotField
        getPivotItem(): excel.PivotItem
        getPivotTable(): excel.PivotTable
        setShowDetail(pShowDetail: any): void
        getListObject(): excel.ListObject
        getAddressLocal(rowAbsolute: any, columnAbsolute: any, referenceStyle: int, external: any, relativeTo: any): string
        isAllowEdit(): boolean
        getColumnWidth(): any
        getCountLarge(): any
        getCurrentArray(): excel.Range
        getCurrentRegion(): excel.Range
        getDependents(): excel.Range
        getPrecedents(): excel.Range
        getDisplayFormat(): excel.DisplayFormat
        getEntireColumn(): excel.Range
        getEntireRow(): excel.Range
        getErrors(): excel.Errors
        getFormulaArray(): any
        getFormulaLabel(): int
        getFormulaLocal(): any
        getFormulaR1C1(): any
        getMergeArea(): excel.Range
        getPageBreak(): int
        getPhonetics(): excel.Phonetics
        getPivotCell(): excel.PivotCell
        getResize(rowSize: any, columnSize: any): excel.Range
        getShowDetail(): any
        getValidation(): excel.Validation
        setValue2(pValue2: any): void
        setFormulaLabel(pFormulaLabel: int): void
        setFormulaLocal(pFormulaLocal: any): void
        setFormulaR1C1(pFormulaR1C1: any): void
        setPageBreak(pPageBreak: int): void
        addComment(text: any): excel.Comment
        advancedFilter(action: int, criteriaRange: any, copyToRange: any, unique: any): any
        applyNames(names: any, ignoreRelativeAbsolute: any, useRowColumnNames: any, omitColumn: any, omitRow: any, order: int, appendLast: any): any
        autoComplete(string: string): string
        autoFilter(field: any, criteria1: any, operator: int, criteria2: any, visibleDropDown: any): any
        paserCriteria(criteria: any): any[]
        autoOutline(): any
        borderAround(lineStyle: any, weight: int, colorIndex: int, color: any, themeColor: any): any
        clearComments(): void
        clearHyperlinks(): void
        clearNotes(): any
        createNames(top: any, left: any, bottom: any, right: any): any
        dataSeries(rowcol: any, type: int, date: int, step: any, stop: any, trend: any): any
        dialogBox(): any
        editionOptions(type: int, option: int, name: any, reference: any, appearance: int, chartSize: int, format: any): any
        fillRight(): any
        findPrevious(after: any): excel.Range
        functionWizard(): any
        insertIndent(insertAmount: int): void
        listNames(): any
        navigateArrow(towardPrecedent: any, arrowNumber: any, linkNumber: any): any
        RemoveDuplicates(columns: any, header: any): void
        rowDifferences(comparison: any): excel.Range
        setPhonetic(): void
        showDependents(remove: any): any
        showErrors(): any
        showPrecedents(remove: any): any
        sortSpecial(sortMethod: int, key1: any, order1: int, type: any, key2: any, order2: int, key3: any, order3: int, header: int, orderCustom: any, matchCase: any, orientation: int, dataOption1: int, dataOption2: int, dataOption3: int): any
        specialCells(type: int, value: any): excel.Range
        textToColumns(destination: any, dataType: int, textQualifier: int, consecutiveDelimiter: any, tab: any, semicolon: any, comma: any, space: any, other: any, otherChar: any, fieldInfo: any, decimalSeparator: any, thousandsSeparator: any, trailingMinusNumbers: any): any
    }
    export interface RangeValue {
        getValue(): any
        setValue(value: any): void
        getType(): int
        setType(type: int): void
        getColumns(): int
        getRows(): int
        getUsedStartColumn(): int
        getUsedStartRow(): int
        getUsedEndRow(): int
        getUsedEndColumn(): int
    }
    export interface RangeValue {
        getValue(): any
        setValue(value: any): void
        getType(): int
        setType(type: int): void
        getColumns(): int
        getRows(): int
        getUsedStartColumn(): int
        getUsedStartRow(): int
        getUsedEndRow(): int
        getUsedEndColumn(): int
    }
    export interface RecentFile {
        getName(): string
        delete(): void
        getPath(): string
        open(): excel.Workbook
        getIndex(): int
    }
    export interface RecentFile {
        getName(): string
        delete(): void
        getPath(): string
        open(): excel.Workbook
        getIndex(): int
    }
    export interface RecentFiles {
        add(name: string): excel.RecentFile
        getCount(): int
        getItem(index: int): excel.RecentFile
        getMaximum(): int
        getMRecentFiles(): any
        setMaximum(pMaximum: int): void
    }
    export interface RecentFiles {
        add(name: string): excel.RecentFile
        getCount(): int
        getItem(index: int): excel.RecentFile
        getMaximum(): int
        getMRecentFiles(): any
        setMaximum(pMaximum: int): void
    }
    export const enum Rectangle {
        //clsid=
    }
    export const enum Rectangle {
        //clsid=
    }
    export interface RoutingSlip {
        getMessage(): any
        reset(): void
        getRecipients(rhs: any): any
        setRecipients(index: any, rhs: any): void
        getSubject(): any
        setSubject(pSubject: any): void
        getDelivery(): int
        isReturnWhenDone(): boolean
        getStatus(): int
        isTrackStatus(): boolean
        setDelivery(pDelivery: int): void
        setMessage(pMessage: any): void
        setTrackStatus(pTrackStatus: boolean): any
        setReturnWhenDone(pReturnWhenDone: boolean): void
    }
    export interface RoutingSlip {
        getMessage(): any
        reset(): void
        getRecipients(rhs: any): any
        setRecipients(index: any, rhs: any): void
        getSubject(): any
        setSubject(pSubject: any): void
        getDelivery(): int
        isReturnWhenDone(): boolean
        getStatus(): int
        isTrackStatus(): boolean
        setDelivery(pDelivery: int): void
        setMessage(pMessage: any): void
        setTrackStatus(pTrackStatus: boolean): any
        setReturnWhenDone(pReturnWhenDone: boolean): void
    }
    export interface RTD {
        getThrottleInterval(): int
        setThrottleInterval(pThrottleInterval: int): void
        refreshData(): void
        restartServers(): void
    }
    export interface RTD {
        getThrottleInterval(): int
        setThrottleInterval(pThrottleInterval: int): void
        refreshData(): void
        restartServers(): void
    }
    export interface Scenario {
        getName(): string
        delete(): any
        setName(pName: string): void
        isHidden(): boolean
        isLocked(): boolean
        getIndex(): int
        getComment(): string
        setComment(pComment: string): void
        show(): any
        setHidden(pHidden: boolean): void
        setLocked(pLocked: boolean): void
        getChangingCells(): excel.Range
        getValues(index: any): any
        changeScenario(changingCells: any, values: any): any
    }
    export interface Scenario {
        getName(): string
        delete(): any
        setName(pName: string): void
        isHidden(): boolean
        isLocked(): boolean
        getIndex(): int
        getComment(): string
        setComment(pComment: string): void
        show(): any
        setHidden(pHidden: boolean): void
        setLocked(pLocked: boolean): void
        getChangingCells(): excel.Range
        getValues(index: any): any
        changeScenario(changingCells: any, values: any): any
    }
    export interface Scenarios {
        add(name: string, changingCells: any, values: any, comment: any, locked: any, hidden: any): excel.Scenario
        merge(source: any): any
        item(index: any): excel.Scenario
        getCount(): int
        createSummary(reportType: int, resultCells: any): any
    }
    export interface Scenarios {
        add(name: string, changingCells: any, values: any, comment: any, locked: any, hidden: any): excel.Scenario
        merge(source: any): any
        item(index: any): excel.Scenario
        getCount(): int
        createSummary(reportType: int, resultCells: any): any
    }
    export interface Series {
        getName(): string
        delete(): any
        setName(pName: string): void
        copy(): any
        getType(): int
        getBorder(): excel.Border
        setType(pType: int): void
        points(index: any): any
        paste(): any
        select(): any
        isShadow(): boolean
        setShadow(pShadow: boolean): void
        getFill(): excel.ChartFillFormat
        clearFormats(): any
        setFormula(pFormula: string): void
        getInterior(): excel.Interior
        setAxisGroup(pAxisGroup: int): void
        getPictureType(): int
        setPictureType(pPictureType: int): void
        setMarkerSize(pMarkerSize: int): void
        getMarkerSize(): int
        setMarkerStyle(pMarkerStyle: int): void
        getMarkerStyle(): int
        setPictureUnit(pPictureUnit: int): void
        getPictureUnit(): int
        isSmooth(): boolean
        errorBar(direction: int, include: int, type: int, amount: any, minusValues: any): any
        setApplyPictToEnd(pApplyPictToEnd: boolean): void
        isApplyPictToFront(): boolean
        setApplyPictToFront(pApplyPictToFront: boolean): void
        isApplyPictToSides(): boolean
        setApplyPictToSides(pApplyPictToSides: boolean): void
        getFormula(): string
        getAxisGroup(): int
        getChartType(): int
        getBarShape(): int
        setBarShape(pBarShape: int): void
        setChartType(pChartType: int): void
        applyCustomType(chartType: int): void
        applyDataLabels(type: int, legendKey: any, autoText: any, hasLeaderLines: any, showSeriesName: any, showCategoryName: any, showValue: any, showPercentage: any, showBubbleSize: any, separator: any): any
        hasLeaderLines(): boolean
        setInvertIfNegative(pInvertIfNegative: boolean): void
        isInvertIfNegative(): boolean
        setMarkerBackgroundColor(pMarkerBackgroundColor: int): void
        getMarkerBackgroundColor(): int
        setMarkerForegroundColor(pMarkerForegroundColor: int): void
        getMarkerForegroundColor(): int
        setMarkerBackgroundColorIndex(pMarkerBackgroundColorIndex: int): void
        getMarkerBackgroundColorIndex(): int
        setMarkerForegroundColorIndex(pMarkerForegroundColorIndex: int): void
        getMarkerForegroundColorIndex(): int
        getFormulaR1C1Local(): string
        setFormulaR1C1Local(pFormulaR1C1Local: string): void
        setHasLeaderLines(pHasLeaderLines: boolean): void
        isApplyPictToEnd(): boolean
        getExplosion(): int
        setExplosion(pExplosion: int): void
        getFormulaLocal(): string
        getFormulaR1C1(): string
        setFormulaLocal(pFormulaLocal: string): void
        setFormulaR1C1(pFormulaR1C1: string): void
        getValues(): any
        getBubbleSizes(): any
        setBubbleSizes(pBubbleSizes: any): void
        getErrorBars(): excel.ErrorBars
        has3DEffect(): boolean
        setHas3DEffect(pHas3DEffect: boolean): void
        hasDataLabels(): boolean
        setHasDataLabels(pHasDataLabels: boolean): void
        hasErrorBars(): boolean
        setHasErrorBars(pHasErrorBars: boolean): void
        getLeaderLines(): excel.LeaderLines
        getPlotOrder(): int
        setPlotOrder(pPlotOrder: int): void
        setSmooth(pSmooth: boolean): void
        setValues(pValues: any): void
        getXValues(): any
        setXValues(pXValues: any): void
        dataLabels(index: any): any
        trendlines(index: any): any
    }
    export interface Series {
        getName(): string
        delete(): any
        setName(pName: string): void
        copy(): any
        getType(): int
        getBorder(): excel.Border
        setType(pType: int): void
        points(index: any): any
        paste(): any
        select(): any
        isShadow(): boolean
        setShadow(pShadow: boolean): void
        getFill(): excel.ChartFillFormat
        clearFormats(): any
        setFormula(pFormula: string): void
        getInterior(): excel.Interior
        setAxisGroup(pAxisGroup: int): void
        getPictureType(): int
        setPictureType(pPictureType: int): void
        setMarkerSize(pMarkerSize: int): void
        getMarkerSize(): int
        setMarkerStyle(pMarkerStyle: int): void
        getMarkerStyle(): int
        setPictureUnit(pPictureUnit: int): void
        getPictureUnit(): int
        isSmooth(): boolean
        errorBar(direction: int, include: int, type: int, amount: any, minusValues: any): any
        setApplyPictToEnd(pApplyPictToEnd: boolean): void
        isApplyPictToFront(): boolean
        setApplyPictToFront(pApplyPictToFront: boolean): void
        isApplyPictToSides(): boolean
        setApplyPictToSides(pApplyPictToSides: boolean): void
        getFormula(): string
        getAxisGroup(): int
        getChartType(): int
        getBarShape(): int
        setBarShape(pBarShape: int): void
        setChartType(pChartType: int): void
        applyCustomType(chartType: int): void
        applyDataLabels(type: int, legendKey: any, autoText: any, hasLeaderLines: any, showSeriesName: any, showCategoryName: any, showValue: any, showPercentage: any, showBubbleSize: any, separator: any): any
        hasLeaderLines(): boolean
        setInvertIfNegative(pInvertIfNegative: boolean): void
        isInvertIfNegative(): boolean
        setMarkerBackgroundColor(pMarkerBackgroundColor: int): void
        getMarkerBackgroundColor(): int
        setMarkerForegroundColor(pMarkerForegroundColor: int): void
        getMarkerForegroundColor(): int
        setMarkerBackgroundColorIndex(pMarkerBackgroundColorIndex: int): void
        getMarkerBackgroundColorIndex(): int
        setMarkerForegroundColorIndex(pMarkerForegroundColorIndex: int): void
        getMarkerForegroundColorIndex(): int
        getFormulaR1C1Local(): string
        setFormulaR1C1Local(pFormulaR1C1Local: string): void
        setHasLeaderLines(pHasLeaderLines: boolean): void
        isApplyPictToEnd(): boolean
        getExplosion(): int
        setExplosion(pExplosion: int): void
        getFormulaLocal(): string
        getFormulaR1C1(): string
        setFormulaLocal(pFormulaLocal: string): void
        setFormulaR1C1(pFormulaR1C1: string): void
        getValues(): any
        getBubbleSizes(): any
        setBubbleSizes(pBubbleSizes: any): void
        getErrorBars(): excel.ErrorBars
        has3DEffect(): boolean
        setHas3DEffect(pHas3DEffect: boolean): void
        hasDataLabels(): boolean
        setHasDataLabels(pHasDataLabels: boolean): void
        hasErrorBars(): boolean
        setHasErrorBars(pHasErrorBars: boolean): void
        getLeaderLines(): excel.LeaderLines
        getPlotOrder(): int
        setPlotOrder(pPlotOrder: int): void
        setSmooth(pSmooth: boolean): void
        setValues(pValues: any): void
        getXValues(): any
        setXValues(pXValues: any): void
        dataLabels(index: any): any
        trendlines(index: any): any
    }
    export interface SeriesCollection {
        add(source: any, rowcol: int, eriesLabels: any, categoryLabels: any, replace: any): excel.Series
        item(index: any): excel.Series
        getCount(): int
        paste(rowcol: int, seriesLabels: any, categoryLabels: any, replace: any, newSeries: any): any
        extend(source: any, rowcol: any, categoryLabels: any): any
        newSeries(): excel.Series
    }
    export interface SeriesCollection {
        add(source: any, rowcol: int, eriesLabels: any, categoryLabels: any, replace: any): excel.Series
        item(index: any): excel.Series
        getCount(): int
        paste(rowcol: int, seriesLabels: any, categoryLabels: any, replace: any, newSeries: any): any
        extend(source: any, rowcol: any, categoryLabels: any): any
        newSeries(): excel.Series
    }
    export interface SeriesLines {
        getName(): string
        delete(): any
        getBorder(): excel.Border
        select(): any
    }
    export interface SeriesLines {
        getName(): string
        delete(): any
        getBorder(): excel.Border
        select(): any
    }
    export interface ShadowFormat {
        getType(): int
        getSize(): double
        setSize(size: double): void
        setType(pType: int): void
        setVisible(pVisible: int): void
        getForeColor(): excel.ColorFormat
        getTransparency(): double
        setTransparency(pTransparency: double): void
        getVisible(): int
        setForeColor(pForeColor: excel.ColorFormat): void
        getObscured(): int
        setObscured(pObscured: int): void
        getOffsetX(): double
        setOffsetX(pOffsetX: double): void
        getOffsetY(): double
        setOffsetY(pOffsetY: double): void
        incrementOffsetX(increment: float): void
        incrementOffsetY(increment: float): void
        getBlur(): double
        setBlur(blur: double): void
    }
    export interface ShadowFormat {
        getType(): int
        getSize(): double
        setSize(size: double): void
        setType(pType: int): void
        setVisible(pVisible: int): void
        getForeColor(): excel.ColorFormat
        getTransparency(): double
        setTransparency(pTransparency: double): void
        getVisible(): int
        setForeColor(pForeColor: excel.ColorFormat): void
        getObscured(): int
        setObscured(pObscured: int): void
        getOffsetX(): double
        setOffsetX(pOffsetX: double): void
        getOffsetY(): double
        setOffsetY(pOffsetY: double): void
        incrementOffsetX(increment: float): void
        incrementOffsetY(increment: float): void
        getBlur(): double
        setBlur(blur: double): void
    }
    export interface Shape {
        getName(): string
        apply(): void
        delete(): void
        setName(pName: string): void
        copy(): void
        getType(): int
        flip(flipCmd: int): void
        getID(): int
        duplicate(): excel.Shape
        getScript(): office.Script
        setTitle(title: string): void
        getTitle(): string
        getWidth(): double
        getLeft(): double
        setLeft(pLeft: double): void
        setTop(pTop: double): void
        getTop(): double
        setWidth(pWidth: double): void
        getHeight(): double
        setHeight(pHeight: double): void
        setVisible(pVisible: int): void
        saveAs(savePath: string): void
        cut(): void
        select(replace: any): void
        getFill(): excel.FillFormat
        getLine(): excel.LineFormat
        getOLEFormat(): excel.OLEFormat
        getHyperlink(): excel.Hyperlink
        getShapeRange(): excel.ShapeRange
        getLocked(): boolean
        getVisible(): int
        getShadow(): excel.ShadowFormat
        getMShape(): any
        getTextFrame(): excel.TextFrame
        getHasChart(): int
        getHasSmartArt(): int
        getPictureFormat(): excel.PictureFormat
        getTextEffect(): excel.TextEffectFormat
        getGroupItems(): excel.GroupShapes
        setAlternativeText(pAlternativeText: string): void
        getAlternativeText(): string
        setLockAspectRatio(pLockAspectRatio: int): void
        getLockAspectRatio(): int
        getLinkFormat(): excel.LinkFormat
        setLocked(pLocked: boolean): void
        ungroup(): excel.ShapeRange
        getNodes(): excel.ShapeNodes
        setShapesDefaultProperties(): void
        getAdjustments(): excel.Adjustments
        getAutoShapeType(): int
        getCallout(): excel.CalloutFormat
        getDiagramNode(): excel.DiagramNode
        getParentGroup(): excel.Shape
        getRotation(): double
        getThreeD(): excel.ThreeDFormat
        getVerticalFlip(): int
        getVertices(): any
        setAutoShapeType(pAutoShapeType: int): void
        setRotation(pRotation: double): void
        incrementLeft(increment: float): void
        incrementTop(increment: float): void
        scaleHeight(factor: double, relativeToOriginalSize: int, scale: any): void
        scaleWidth(factor: double, relativeToOriginalSize: int, scale: any): void
        setOthersMarkID(markID: string): void
        getOthersMarkID(): string
        getHorizontalFlip(): int
        getZOrderPosition(): int
        incrementRotation(increment: float): void
        setOthersPressValues(pressValues: number[]): void
        getOthersPressValue(): number[]
        setOthersUserName(name: string): void
        getOthersUserName(): string
        setOthersCreateTime(time: string): void
        getOthersCreateTime(): string
        getChild(): int
        pickUp(): void
        zOrder(zOrderCmd: int): void
        isIsf(): boolean
        setShapeStyle(style: int): void
        getConnector(): int
        getChart(): excel.Chart
        copyPicture(appearance: any, format: any): void
        getTopLeftCell(): excel.Range
        setPlacement(pPlacement: int): void
        getPlacement(): int
        getControlFormat(): excel.ControlFormat
        getOnAction(): string
        setOnAction(pOnAction: string): void
        getTextFrame2(): excel.TextFrame2
        getShapeStyle(): int
        getBottomRightCell(): excel.Range
        getBlackWhiteMode(): int
        setBlackWhiteMode(pBlackWhiteMode: int): void
        getConnectionSiteCount(): int
        getConnectorFormat(): excel.ConnectorFormat
        getFormControlType(): int
        rerouteConnections(): void
    }
    export interface Shape {
        getName(): string
        apply(): void
        delete(): void
        setName(pName: string): void
        copy(): void
        getType(): int
        flip(flipCmd: int): void
        getID(): int
        duplicate(): excel.Shape
        getScript(): office.Script
        setTitle(title: string): void
        getTitle(): string
        getWidth(): double
        getLeft(): double
        setLeft(pLeft: double): void
        setTop(pTop: double): void
        getTop(): double
        setWidth(pWidth: double): void
        getHeight(): double
        setHeight(pHeight: double): void
        setVisible(pVisible: int): void
        saveAs(savePath: string): void
        cut(): void
        select(replace: any): void
        getFill(): excel.FillFormat
        getLine(): excel.LineFormat
        getOLEFormat(): excel.OLEFormat
        getHyperlink(): excel.Hyperlink
        getShapeRange(): excel.ShapeRange
        getLocked(): boolean
        getVisible(): int
        getShadow(): excel.ShadowFormat
        getMShape(): any
        getTextFrame(): excel.TextFrame
        getHasChart(): int
        getHasSmartArt(): int
        getPictureFormat(): excel.PictureFormat
        getTextEffect(): excel.TextEffectFormat
        getGroupItems(): excel.GroupShapes
        setAlternativeText(pAlternativeText: string): void
        getAlternativeText(): string
        setLockAspectRatio(pLockAspectRatio: int): void
        getLockAspectRatio(): int
        getLinkFormat(): excel.LinkFormat
        setLocked(pLocked: boolean): void
        ungroup(): excel.ShapeRange
        getNodes(): excel.ShapeNodes
        setShapesDefaultProperties(): void
        getAdjustments(): excel.Adjustments
        getAutoShapeType(): int
        getCallout(): excel.CalloutFormat
        getDiagramNode(): excel.DiagramNode
        getParentGroup(): excel.Shape
        getRotation(): double
        getThreeD(): excel.ThreeDFormat
        getVerticalFlip(): int
        getVertices(): any
        setAutoShapeType(pAutoShapeType: int): void
        setRotation(pRotation: double): void
        incrementLeft(increment: float): void
        incrementTop(increment: float): void
        scaleHeight(factor: double, relativeToOriginalSize: int, scale: any): void
        scaleWidth(factor: double, relativeToOriginalSize: int, scale: any): void
        setOthersMarkID(markID: string): void
        getOthersMarkID(): string
        getHorizontalFlip(): int
        getZOrderPosition(): int
        incrementRotation(increment: float): void
        setOthersPressValues(pressValues: number[]): void
        getOthersPressValue(): number[]
        setOthersUserName(name: string): void
        getOthersUserName(): string
        setOthersCreateTime(time: string): void
        getOthersCreateTime(): string
        getChild(): int
        pickUp(): void
        zOrder(zOrderCmd: int): void
        isIsf(): boolean
        setShapeStyle(style: int): void
        getConnector(): int
        getChart(): excel.Chart
        copyPicture(appearance: any, format: any): void
        getTopLeftCell(): excel.Range
        setPlacement(pPlacement: int): void
        getPlacement(): int
        getControlFormat(): excel.ControlFormat
        getOnAction(): string
        setOnAction(pOnAction: string): void
        getTextFrame2(): excel.TextFrame2
        getShapeStyle(): int
        getBottomRightCell(): excel.Range
        getBlackWhiteMode(): int
        setBlackWhiteMode(pBlackWhiteMode: int): void
        getConnectionSiteCount(): int
        getConnectorFormat(): excel.ConnectorFormat
        getFormControlType(): int
        rerouteConnections(): void
    }
    export interface ShapeNode {
        getEditingType(): int
        getPoints(): any
        getSegmentType(): int
    }
    export interface ShapeNode {
        getEditingType(): int
        getPoints(): any
        getSegmentType(): int
    }
    export interface ShapeNodes {
        delete(index: int): void
        insert(index: int, segmentType: int, editingType: int, x1: double, y1: double, x2: double, y2: double, x3: double, y3: double): void
        item(index: any): excel.ShapeNode
        getCount(): int
        setPosition(index: int, x1: double, y1: double): void
        setEditingType(index: int, editingType: int): void
        setSegmentType(index: int, segmentType: int): void
    }
    export interface ShapeNodes {
        delete(index: int): void
        insert(index: int, segmentType: int, editingType: int, x1: double, y1: double, x2: double, y2: double, x3: double, y3: double): void
        item(index: any): excel.ShapeNode
        getCount(): int
        setPosition(index: int, x1: double, y1: double): void
        setEditingType(index: int, editingType: int): void
        setSegmentType(index: int, segmentType: int): void
    }
    export interface ShapeRange {
        group(): excel.Shape
        getName(): string
        apply(): void
        delete(): void
        setName(pName: string): void
        getType(): int
        flip(flipCmd: int): void
        getID(): int
        duplicate(): excel.ShapeRange
        item(index: any): excel.Shape
        getCount(): int
        setTitle(title: string): void
        getTitle(): string
        getWidth(): double
        getLeft(): double
        setLeft(pLeft: double): void
        setTop(pTop: double): void
        getTop(): double
        setWidth(pWidth: double): void
        getHeight(): double
        setHeight(pHeight: double): void
        setVisible(pVisible: int): void
        select(replace: any): void
        align(alignCmd: int, relativeTo: int): void
        getFill(): excel.FillFormat
        getLine(): excel.LineFormat
        getVisible(): int
        getShadow(): excel.ShadowFormat
        getTextFrame(): excel.TextFrame
        getHasChart(): int
        getPictureFormat(): excel.PictureFormat
        getTextEffect(): excel.TextEffectFormat
        getGroupItems(): excel.GroupShapes
        setAlternativeText(pAlternativeText: string): void
        getAlternativeText(): string
        setLockAspectRatio(pLockAspectRatio: int): void
        getLockAspectRatio(): int
        getDiagram(): excel.Diagram
        ungroup(): excel.ShapeRange
        getNodes(): excel.ShapeNodes
        setShapesDefaultProperties(): void
        regroup(): excel.Shape
        getAdjustments(): excel.Adjustments
        getAutoShapeType(): int
        getCallout(): excel.CalloutFormat
        getDiagramNode(): excel.DiagramNode
        getHasDiagram(): int
        getParentGroup(): excel.Shape
        getRotation(): double
        getThreeD(): excel.ThreeDFormat
        getVerticalFlip(): int
        getVertices(): any
        setAutoShapeType(pAutoShapeType: int): void
        setRotation(pRotation: double): void
        incrementLeft(increment: double): void
        incrementTop(increment: double): void
        scaleHeight(factor: double, relativeToOriginalSize: int, scale: any): void
        scaleWidth(factor: double, relativeToOriginalSize: int, scale: any): void
        distribute(distributeCmd: int, relativeTo: int): void
        getHasDiagramNode(): int
        getHorizontalFlip(): int
        getZOrderPosition(): int
        incrementRotation(increment: double): void
        getChild(): int
        pickUp(): void
        zOrder(zOrderCmd: int): void
        setShapeStyle(style: int): void
        getConnector(): int
        getBlackWhiteMode(): int
        setBlackWhiteMode(pBlackWhiteMode: int): void
        getConnectionSiteCount(): int
        getConnectorFormat(): excel.ConnectorFormat
        rerouteConnections(): void
    }
    export interface ShapeRange {
        group(): excel.Shape
        getName(): string
        apply(): void
        delete(): void
        setName(pName: string): void
        getType(): int
        flip(flipCmd: int): void
        getID(): int
        duplicate(): excel.ShapeRange
        item(index: any): excel.Shape
        getCount(): int
        setTitle(title: string): void
        getTitle(): string
        getWidth(): double
        getLeft(): double
        setLeft(pLeft: double): void
        setTop(pTop: double): void
        getTop(): double
        setWidth(pWidth: double): void
        getHeight(): double
        setHeight(pHeight: double): void
        setVisible(pVisible: int): void
        select(replace: any): void
        align(alignCmd: int, relativeTo: int): void
        getFill(): excel.FillFormat
        getLine(): excel.LineFormat
        getVisible(): int
        getShadow(): excel.ShadowFormat
        getTextFrame(): excel.TextFrame
        getHasChart(): int
        getPictureFormat(): excel.PictureFormat
        getTextEffect(): excel.TextEffectFormat
        getGroupItems(): excel.GroupShapes
        setAlternativeText(pAlternativeText: string): void
        getAlternativeText(): string
        setLockAspectRatio(pLockAspectRatio: int): void
        getLockAspectRatio(): int
        getDiagram(): excel.Diagram
        ungroup(): excel.ShapeRange
        getNodes(): excel.ShapeNodes
        setShapesDefaultProperties(): void
        regroup(): excel.Shape
        getAdjustments(): excel.Adjustments
        getAutoShapeType(): int
        getCallout(): excel.CalloutFormat
        getDiagramNode(): excel.DiagramNode
        getHasDiagram(): int
        getParentGroup(): excel.Shape
        getRotation(): double
        getThreeD(): excel.ThreeDFormat
        getVerticalFlip(): int
        getVertices(): any
        setAutoShapeType(pAutoShapeType: int): void
        setRotation(pRotation: double): void
        incrementLeft(increment: double): void
        incrementTop(increment: double): void
        scaleHeight(factor: double, relativeToOriginalSize: int, scale: any): void
        scaleWidth(factor: double, relativeToOriginalSize: int, scale: any): void
        distribute(distributeCmd: int, relativeTo: int): void
        getHasDiagramNode(): int
        getHorizontalFlip(): int
        getZOrderPosition(): int
        incrementRotation(increment: double): void
        getChild(): int
        pickUp(): void
        zOrder(zOrderCmd: int): void
        setShapeStyle(style: int): void
        getConnector(): int
        getBlackWhiteMode(): int
        setBlackWhiteMode(pBlackWhiteMode: int): void
        getConnectionSiteCount(): int
        getConnectorFormat(): excel.ConnectorFormat
        rerouteConnections(): void
    }
    export interface Shapes {
        item(index: any): excel.Shape
        getShape(index: any): excel.Shape
        getCount(): int
        getRange(index: any): excel.ShapeRange
        addCallout(type: int, left: double, top: double, width: double, height: double): excel.Shape
        addConnector(type: int, beginX: double, beginY: double, endX: double, endY: double): excel.Shape
        addPicture(filename: string, linkToFile: int, saveWithDocument: int, left: double, top: double, width: double, height: double): excel.Shape
        addPolyline(safeArrayOfPoints: any): excel.Shape
        addTextbox(orientation: int, left: double, top: double, width: double, height: double): excel.Shape
        addTextEffect(presetTextEffect: int, text: string, fontName: string, fontSize: double, fontBold: int, fontItalic: int, left: double, top: double): excel.Shape
        buildFreeform(editingType: int, x1: double, y1: double): excel.FreeformBuilder
        selectAll(): void
        addCurve(safeArrayOfPoints: any): excel.Shape
        addLabel(orientation: int, left: double, top: double, width: double, height: double): excel.Shape
        addLine(beginX: double, beginY: double, endX: double, endY: double): excel.Shape
        addShape(type: int, left: double, top: double, width: double, height: double): excel.Shape
        addOLEObject(classType: any, fileName: any, linkToFile: any, displayAsIcon: any, iconFileName: any, iconIndex: any, iconLabel: any, left: any, top: any, width: any, height: any): excel.Shape
        getAllShapes(): excel.Shape[]
        addDiagram(type: int, left: float, top: float, width: float, height: float): excel.Shape
        addChart(type: int, left: float, top: float, width: float, height: float): excel.Shape
        addFormControl(type: int, left: int, top: int, width: int, height: int): excel.Shape
    }
    export interface Shapes {
        item(index: any): excel.Shape
        getShape(index: any): excel.Shape
        getCount(): int
        getRange(index: any): excel.ShapeRange
        addCallout(type: int, left: double, top: double, width: double, height: double): excel.Shape
        addConnector(type: int, beginX: double, beginY: double, endX: double, endY: double): excel.Shape
        addPicture(filename: string, linkToFile: int, saveWithDocument: int, left: double, top: double, width: double, height: double): excel.Shape
        addPolyline(safeArrayOfPoints: any): excel.Shape
        addTextbox(orientation: int, left: double, top: double, width: double, height: double): excel.Shape
        addTextEffect(presetTextEffect: int, text: string, fontName: string, fontSize: double, fontBold: int, fontItalic: int, left: double, top: double): excel.Shape
        buildFreeform(editingType: int, x1: double, y1: double): excel.FreeformBuilder
        selectAll(): void
        addCurve(safeArrayOfPoints: any): excel.Shape
        addLabel(orientation: int, left: double, top: double, width: double, height: double): excel.Shape
        addLine(beginX: double, beginY: double, endX: double, endY: double): excel.Shape
        addShape(type: int, left: double, top: double, width: double, height: double): excel.Shape
        addOLEObject(classType: any, fileName: any, linkToFile: any, displayAsIcon: any, iconFileName: any, iconIndex: any, iconLabel: any, left: any, top: any, width: any, height: any): excel.Shape
        getAllShapes(): excel.Shape[]
        addDiagram(type: int, left: float, top: float, width: float, height: float): excel.Shape
        addChart(type: int, left: float, top: float, width: float, height: float): excel.Shape
        addFormControl(type: int, left: int, top: int, width: int, height: int): excel.Shape
    }
    export interface Sheets {
        add(before: any, after: any, count: any, type: any): any
        delete(): void
        item(index: any): any
        getCount(): int
        printPreview(enableChanges: any): void
        setVisible(pVisible: any): void
        printOut(from: any, to: any, copies: any, preview: any, activePrinter: any, printToFile: any, collate: any, prToFileName: any): void
        printOut(from: any, to: any, copies: any, preview: any, activePrinter: any, printToFile: any, collate: any, prToFileName: any, ignorePrintAreas: any): void
        select(replace: any): void
        getVisible(): any
        getHPageBreaks(): excel.HPageBreaks
        getVPageBreaks(): excel.VPageBreaks
        fillAcrossSheets(range: excel.Range, type: int): void
    }
    export interface Sheets {
        add(before: any, after: any, count: any, type: any): any
        delete(): void
        item(index: any): any
        getCount(): int
        printPreview(enableChanges: any): void
        setVisible(pVisible: any): void
        printOut(from: any, to: any, copies: any, preview: any, activePrinter: any, printToFile: any, collate: any, prToFileName: any): void
        printOut(from: any, to: any, copies: any, preview: any, activePrinter: any, printToFile: any, collate: any, prToFileName: any, ignorePrintAreas: any): void
        select(replace: any): void
        getVisible(): any
        getHPageBreaks(): excel.HPageBreaks
        getVPageBreaks(): excel.VPageBreaks
        fillAcrossSheets(range: excel.Range, type: int): void
    }
    export interface SheetViews {
        item(index: any): excel.WorksheetView
        getCount(): int
    }
    export interface SheetViews {
        item(index: any): excel.WorksheetView
        getCount(): int
    }
    export interface SmartTag {
        getName(): string
        getProperties(): excel.CustomProperties
        delete(): void
        getRange(): excel.Range
        getXML(): string
        getDownloadURL(): string
        getSmartTagActions(): excel.SmartTagActions
    }
    export interface SmartTag {
        getName(): string
        getProperties(): excel.CustomProperties
        delete(): void
        getRange(): excel.Range
        getXML(): string
        getDownloadURL(): string
        getSmartTagActions(): excel.SmartTagActions
    }
    export interface SmartTagAction {
        getName(): string
        execute(): void
        getType(): int
        isCheckboxState(): boolean
        isExpandHelp(): boolean
        getListSelection(): int
        isPresentInPane(): boolean
        getTextboxText(): string
        setCheckboxState(pCheckboxState: boolean): void
        setExpandHelp(pExpandHelp: boolean): void
        setListSelection(pListSelection: int): void
        setTextboxText(pTextboxText: string): void
        getActiveXControl(): any
        getRadioGroupSelection(): int
        setRadioGroupSelection(pRadioGroupSelection: int): void
    }
    export interface SmartTagAction {
        getName(): string
        execute(): void
        getType(): int
        isCheckboxState(): boolean
        isExpandHelp(): boolean
        getListSelection(): int
        isPresentInPane(): boolean
        getTextboxText(): string
        setCheckboxState(pCheckboxState: boolean): void
        setExpandHelp(pExpandHelp: boolean): void
        setListSelection(pListSelection: int): void
        setTextboxText(pTextboxText: string): void
        getActiveXControl(): any
        getRadioGroupSelection(): int
        setRadioGroupSelection(pRadioGroupSelection: int): void
    }
    export interface SmartTagActions {
        item(index: any): excel.SmartTagAction
        getCount(): int
    }
    export interface SmartTagActions {
        item(index: any): excel.SmartTagAction
        getCount(): int
    }
    export interface SmartTagOptions {
        setEmbedSmartTags(pEmbedSmartTags: boolean): void
        isEmbedSmartTags(): boolean
        setDisplaySmartTags(pDisplaySmartTags: int): void
        getDisplaySmartTags(): int
    }
    export interface SmartTagOptions {
        setEmbedSmartTags(pEmbedSmartTags: boolean): void
        isEmbedSmartTags(): boolean
        setDisplaySmartTags(pDisplaySmartTags: int): void
        getDisplaySmartTags(): int
    }
    export interface SmartTagRecognizer {
        getFullName(): string
        setEnabled(pEnabled: boolean): void
        isEnabled(): boolean
        getprogID(): string
    }
    export interface SmartTagRecognizer {
        getFullName(): string
        setEnabled(pEnabled: boolean): void
        isEnabled(): boolean
        getprogID(): string
    }
    export interface SmartTagRecognizers {
        item(index: any): excel.SmartTagRecognizer
        getCount(): int
        isRecognize(): boolean
        setRecognize(pRecognize: boolean): void
    }
    export interface SmartTagRecognizers {
        item(index: any): excel.SmartTagRecognizer
        getCount(): int
        isRecognize(): boolean
        setRecognize(pRecognize: boolean): void
    }
    export interface SmartTags {
        add(smartTagType: string): excel.SmartTag
        getCount(): int
    }
    export interface SmartTags {
        add(smartTagType: string): excel.SmartTag
        getCount(): int
    }
    export interface Sort {
        apply(): void
        getOrientation(): int
        setOrientation(orientation: int): void
        setMatchCase(match: boolean): void
        getMatchCase(): boolean
        getHeader(): int
        setRange(range: excel.Range): void
        setHeader(header: int): void
        getRng(): excel.Range
        getSortFields(): excel.SortFields
        getSortMethod(): int
        setSortMethod(stroke: int): void
    }
    export interface Sort {
        apply(): void
        getOrientation(): int
        setOrientation(orientation: int): void
        setMatchCase(match: boolean): void
        getMatchCase(): boolean
        getHeader(): int
        setRange(range: excel.Range): void
        setHeader(header: int): void
        getRng(): excel.Range
        getSortFields(): excel.SortFields
        getSortMethod(): int
        setSortMethod(stroke: int): void
    }
    export interface SortField {
        getKey(): excel.Range
        delete(): void
        setPriority(priority: int): void
        getPriority(): int
        setIcon(icon: excel.Icon): void
        getOrder(): int
        setOrder(order: int): void
        getCustomOrder(): any
        setCustomOrder(order: any): void
        getDataOption(): int
        setDataOption(dataOption: int): void
        getSortOn(): int
        setSortOn(sortOn: int): void
        getSortOnValue(): any
        modifyKey(key: excel.Range): void
    }
    export interface SortField {
        getKey(): excel.Range
        delete(): void
        setPriority(priority: int): void
        getPriority(): int
        setIcon(icon: excel.Icon): void
        getOrder(): int
        setOrder(order: int): void
        getCustomOrder(): any
        setCustomOrder(order: any): void
        getDataOption(): int
        setDataOption(dataOption: int): void
        getSortOn(): int
        setSortOn(sortOn: int): void
        getSortOnValue(): any
        modifyKey(key: excel.Range): void
    }
    export interface SortFields {
        add(key: excel.Range, sortOn: any, order: any, customOrder: any, dataOption: any): excel.SortField
        remove(sf: excel.SortField): void
        clear(): void
        getCount(): int
        getItem(i: any): excel.SortField
    }
    export interface SortFields {
        add(key: excel.Range, sortOn: any, order: any, customOrder: any, dataOption: any): excel.SortField
        remove(sf: excel.SortField): void
        clear(): void
        getCount(): int
        getItem(i: any): excel.SortField
    }
    export interface Speech {
        getDirection(): int
        setDirection(pDirection: int): void
        isSpeakCellOnEnter(): boolean
        setSpeakCellOnEnter(pSpeakCellOnEnter: boolean): void
        speak(text: string, speakAsync: any, speakXML: boolean, purge: any): void
    }
    export interface Speech {
        getDirection(): int
        setDirection(pDirection: int): void
        isSpeakCellOnEnter(): boolean
        setSpeakCellOnEnter(pSpeakCellOnEnter: boolean): void
        speak(text: string, speakAsync: any, speakXML: boolean, purge: any): void
    }
    export interface SpellingOptions {
        isIgnoreMixedDigits(): boolean
        setIgnoreMixedDigits(pIgnoreMixedDigits: boolean): void
        isKoreanUseAutoChangeList(): boolean
        setKoreanUseAutoChangeList(pKoreanUseAutoChangeList: boolean): void
        isGermanPostReform(): boolean
        setGermanPostReform(pGermanPostReform: boolean): void
        isIgnoreFileNames(): boolean
        setIgnoreFileNames(pIgnoreFileNames: boolean): void
        isKoreanCombineAux(): boolean
        setKoreanCombineAux(pKoreanCombineAux: boolean): void
        isKoreanProcessCompound(): boolean
        setKoreanProcessCompound(pKoreanProcessCompound: boolean): void
        isSuggestMainOnly(): boolean
        setSuggestMainOnly(pSuggestMainOnly: boolean): void
        getArabicModes(): int
        setArabicModes(pArabicModes: int): void
        getDictLang(): int
        setDictLang(pDictLang: int): void
        getHebrewModes(): int
        setHebrewModes(pHebrewModes: int): void
        isIgnoreCaps(): boolean
        setIgnoreCaps(pIgnoreCaps: boolean): void
        getUserDict(): string
        setUserDict(pUserDict: string): void
    }
    export interface SpellingOptions {
        isIgnoreMixedDigits(): boolean
        setIgnoreMixedDigits(pIgnoreMixedDigits: boolean): void
        isKoreanUseAutoChangeList(): boolean
        setKoreanUseAutoChangeList(pKoreanUseAutoChangeList: boolean): void
        isGermanPostReform(): boolean
        setGermanPostReform(pGermanPostReform: boolean): void
        isIgnoreFileNames(): boolean
        setIgnoreFileNames(pIgnoreFileNames: boolean): void
        isKoreanCombineAux(): boolean
        setKoreanCombineAux(pKoreanCombineAux: boolean): void
        isKoreanProcessCompound(): boolean
        setKoreanProcessCompound(pKoreanProcessCompound: boolean): void
        isSuggestMainOnly(): boolean
        setSuggestMainOnly(pSuggestMainOnly: boolean): void
        getArabicModes(): int
        setArabicModes(pArabicModes: int): void
        getDictLang(): int
        setDictLang(pDictLang: int): void
        getHebrewModes(): int
        setHebrewModes(pHebrewModes: int): void
        isIgnoreCaps(): boolean
        setIgnoreCaps(pIgnoreCaps: boolean): void
        getUserDict(): string
        setUserDict(pUserDict: string): void
    }
    export interface Style {
        getName(): string
        getValue(): string
        delete(): any
        setName(name: string): void
        getSize(): float
        setSize(f: float): void
        setColor(color: any, obj: any): void
        setColor(color: any): void
        isLocked(): boolean
        getFont(): excel.Font
        getOrientation(): int
        getColor(): any
        getColor(obj: any): any
        setLineStyle(style: int, width: float, obj: any): void
        getLineStyle(obj: any): int
        getLineWidth(obj: any): float
        getBorders(): excel.Borders
        setVerticalAlignment(pVerticalAlignment: int): void
        getVerticalAlignment(): int
        setBold(l: int): void
        getBold(): int
        setOrientation(pOrientation: int): void
        setItalic(l: int): void
        getItalic(): int
        getNameLocal(): string
        getNumberFormat(): string
        setNumberFormat(pNumberFormat: string): void
        setStyle(flag: boolean, attr: int): void
        setLocked(pLocked: boolean): void
        getReadingOrder(): int
        setReadingOrder(pReadingOrder: int): void
        getMStyle(): any
        isBuiltIn(): boolean
        getHorizontalAlignment(): int
        setHorizontalAlignment(pHorizontalAlignment: int): void
        getInterior(): excel.Interior
        isIncludeAlignment(): boolean
        setIncludeAlignment(pIncludeAlignment: boolean): void
        isIncludePatterns(): boolean
        setIncludePatterns(pIncludePatterns: boolean): void
        isIncludeProtection(): boolean
        setIncludeProtection(pIncludeProtection: boolean): void
        initStyle(): void
        isIncludeBorder(): boolean
        setIncludeBorder(pIncludeBorder: boolean): void
        isIncludeFont(): boolean
        setIncludeFont(pIncludeFont: boolean): void
        isIncludeNumber(): boolean
        setIncludeNumber(pIncludeNumber: boolean): void
        getMergeCells(): any
        setWrapText(pWrapText: boolean): void
        isAddIndent(): boolean
        setAddIndent(pAddIndent: boolean): void
        isFormulaHidden(): boolean
        setFormulaHidden(pFormulaHidden: boolean): void
        setIndentLevel(pIndentLevel: int): void
        getIndentLevel(): int
        setMergeCells(pMergeCells: any): void
        setShrinkToFit(pShrinkToFit: boolean): void
        isShrinkToFit(): boolean
        isWrapText(): boolean
        setNumberFormatLocal(pNumberFormatLocal: string): void
        getNumberFormatLocal(): string
    }
    export interface Style {
        getName(): string
        getValue(): string
        delete(): any
        setName(name: string): void
        getSize(): float
        setSize(f: float): void
        setColor(color: any, obj: any): void
        setColor(color: any): void
        isLocked(): boolean
        getFont(): excel.Font
        getOrientation(): int
        getColor(): any
        getColor(obj: any): any
        setLineStyle(style: int, width: float, obj: any): void
        getLineStyle(obj: any): int
        getLineWidth(obj: any): float
        getBorders(): excel.Borders
        setVerticalAlignment(pVerticalAlignment: int): void
        getVerticalAlignment(): int
        setBold(l: int): void
        getBold(): int
        setOrientation(pOrientation: int): void
        setItalic(l: int): void
        getItalic(): int
        getNameLocal(): string
        getNumberFormat(): string
        setNumberFormat(pNumberFormat: string): void
        setStyle(flag: boolean, attr: int): void
        setLocked(pLocked: boolean): void
        getReadingOrder(): int
        setReadingOrder(pReadingOrder: int): void
        getMStyle(): any
        isBuiltIn(): boolean
        getHorizontalAlignment(): int
        setHorizontalAlignment(pHorizontalAlignment: int): void
        getInterior(): excel.Interior
        isIncludeAlignment(): boolean
        setIncludeAlignment(pIncludeAlignment: boolean): void
        isIncludePatterns(): boolean
        setIncludePatterns(pIncludePatterns: boolean): void
        isIncludeProtection(): boolean
        setIncludeProtection(pIncludeProtection: boolean): void
        initStyle(): void
        isIncludeBorder(): boolean
        setIncludeBorder(pIncludeBorder: boolean): void
        isIncludeFont(): boolean
        setIncludeFont(pIncludeFont: boolean): void
        isIncludeNumber(): boolean
        setIncludeNumber(pIncludeNumber: boolean): void
        getMergeCells(): any
        setWrapText(pWrapText: boolean): void
        isAddIndent(): boolean
        setAddIndent(pAddIndent: boolean): void
        isFormulaHidden(): boolean
        setFormulaHidden(pFormulaHidden: boolean): void
        setIndentLevel(pIndentLevel: int): void
        getIndentLevel(): int
        setMergeCells(pMergeCells: any): void
        setShrinkToFit(pShrinkToFit: boolean): void
        isShrinkToFit(): boolean
        isWrapText(): boolean
        setNumberFormatLocal(pNumberFormatLocal: string): void
        getNumberFormatLocal(): string
    }
    export interface Styles {
        add(name: string, basedOn: any): excel.Style
        merge(workbook: any): void
        item(index: any): excel.Style
        getCount(): int
    }
    export interface Styles {
        add(name: string, basedOn: any): excel.Style
        merge(workbook: any): void
        item(index: any): excel.Style
        getCount(): int
    }
    export interface Tab {
        setColor(pColor: any): void
        getColor(): any
        setColorIndex(pColorIndex: int): void
        getColorIndex(): int
    }
    export interface Tab {
        setColor(pColor: any): void
        getColor(): any
        setColorIndex(pColorIndex: int): void
        getColorIndex(): int
    }
    export const enum TextBox {
        //clsid=
    }
    export const enum TextBox {
        //clsid=
    }
    export interface TextEffectFormat {
        getText(): string
        setFontSize(pFontSize: double): void
        setText(pText: string): void
        getAlignment(): int
        setAlignment(align: int): void
        getFontName(): string
        setFontName(pFontName: string): void
        setNormalizedHeight(pNormalizedHeight: int): void
        getNormalizedHeight(): int
        setPresetTextEffect(pPresetTextEffect: int): void
        getPresetTextEffect(): int
        toggleVerticalText(): void
        setFontBold(pFontBold: int): void
        getFontBold(): int
        setFontItalic(pFontItalic: int): void
        getFontItalic(): int
        getFontSize(): double
        setKernedPairs(pKernedPairs: int): void
        setPresetShape(pPresetShape: int): void
        getPresetShape(): int
        setRotatedChars(pRotatedChars: int): void
        getRotatedChars(): int
        setTracking(pTracking: double): void
        getTracking(): double
        getKernedPairs(): int
    }
    export interface TextEffectFormat {
        getText(): string
        setFontSize(pFontSize: double): void
        setText(pText: string): void
        getAlignment(): int
        setAlignment(align: int): void
        getFontName(): string
        setFontName(pFontName: string): void
        setNormalizedHeight(pNormalizedHeight: int): void
        getNormalizedHeight(): int
        setPresetTextEffect(pPresetTextEffect: int): void
        getPresetTextEffect(): int
        toggleVerticalText(): void
        setFontBold(pFontBold: int): void
        getFontBold(): int
        setFontItalic(pFontItalic: int): void
        getFontItalic(): int
        getFontSize(): double
        setKernedPairs(pKernedPairs: int): void
        setPresetShape(pPresetShape: int): void
        getPresetShape(): int
        setRotatedChars(pRotatedChars: int): void
        getRotatedChars(): int
        setTracking(pTracking: double): void
        getTracking(): double
        getKernedPairs(): int
    }
    export interface TextFrame {
        getOrientation(): int
        setAutoSize(b: boolean): void
        isAutoSize(): boolean
        setVerticalAlignment(v: int): void
        getVerticalAlignment(): int
        setOrientation(ori: int): void
        getReadingOrder(): int
        setReadingOrder(pReadingOrder: int): void
        getMarginLeft(): double
        setMarginLeft(value: double): void
        getMarginRight(): double
        setMarginRight(value: double): void
        getMarginBottom(): double
        setMarginBottom(value: double): void
        getMarginTop(): double
        setMarginTop(value: double): void
        getHorizontalAlignment(): int
        setHorizontalAlignment(v: int): void
        isAutoMargins(): boolean
        setAutoMargins(pAutoMargins: boolean): void
        characters(start: any, length: any): excel.Characters
    }
    export interface TextFrame {
        getOrientation(): int
        setAutoSize(b: boolean): void
        isAutoSize(): boolean
        setVerticalAlignment(v: int): void
        getVerticalAlignment(): int
        setOrientation(ori: int): void
        getReadingOrder(): int
        setReadingOrder(pReadingOrder: int): void
        getMarginLeft(): double
        setMarginLeft(value: double): void
        getMarginRight(): double
        setMarginRight(value: double): void
        getMarginBottom(): double
        setMarginBottom(value: double): void
        getMarginTop(): double
        setMarginTop(value: double): void
        getHorizontalAlignment(): int
        setHorizontalAlignment(v: int): void
        isAutoMargins(): boolean
        setAutoMargins(pAutoMargins: boolean): void
        characters(start: any, length: any): excel.Characters
    }
    export interface TextFrame2 {
        getOrientation(): int
        setAutoSize(v: int): void
        setWordWrap(wr: int): void
        setOrientation(ori: int): void
        getTextRange(): office.TextRange2
        getWordWrap(): int
        hasText(): int
        getThreeD(): excel.ThreeDFormat
        getMarginLeft(): double
        setMarginLeft(value: double): void
        getMarginRight(): double
        setMarginRight(value: double): void
        getMarginBottom(): double
        setMarginBottom(value: double): void
        getMarginTop(): double
        setMarginTop(value: double): void
        setHorizontalAnchor(v: int): void
        getHorizontalAnchor(): int
        setVerticalAnchor(v: int): void
        getVerticalAnchor(): int
        getAutoSize(): int
        deleteText(): void
        getNoTextRotation(): int
        setNoTextRotation(v: int): void
        setWordArtformat(wr: int): void
        getWordArtformat(): int
    }
    export interface TextFrame2 {
        getOrientation(): int
        setAutoSize(v: int): void
        setWordWrap(wr: int): void
        setOrientation(ori: int): void
        getTextRange(): office.TextRange2
        getWordWrap(): int
        hasText(): int
        getThreeD(): excel.ThreeDFormat
        getMarginLeft(): double
        setMarginLeft(value: double): void
        getMarginRight(): double
        setMarginRight(value: double): void
        getMarginBottom(): double
        setMarginBottom(value: double): void
        getMarginTop(): double
        setMarginTop(value: double): void
        setHorizontalAnchor(v: int): void
        getHorizontalAnchor(): int
        setVerticalAnchor(v: int): void
        getVerticalAnchor(): int
        getAutoSize(): int
        deleteText(): void
        getNoTextRotation(): int
        setNoTextRotation(v: int): void
        setWordArtformat(wr: int): void
        getWordArtformat(): int
    }
    export interface ThreeDFormat {
        setVisible(pVisible: int): void
        getVisible(): int
        setDepth(pDepth: double): void
        getDepth(): double
        getExtrusionColor(): excel.ColorFormat
        setExtrusionColorType(pExtrusionColorType: int): void
        getExtrusionColorType(): int
        setPresetMaterial(pPresetMaterial: int): void
        getPresetMaterial(): int
        incrementRotationX(increment: double): void
        incrementRotationY(increment: double): void
        getPresetThreeDFormat(): int
        getPresetExtrusionDirection(): int
        setPresetLightingDirection(pPresetLightingDirection: int): void
        getPresetLightingDirection(): int
        setPresetLightingSoftness(pPresetLightingSoftness: int): void
        getPresetLightingSoftness(): int
        setPerspective(pPerspective: int): void
        getPerspective(): int
        setRotationX(pRotationX: double): void
        getRotationX(): double
        setRotationY(pRotationY: double): void
        getRotationY(): double
        resetRotation(): void
        setExtrusionDirection(presetExtrusionDirection: int): void
        setPresetThreeDFormat(presetThreeDFormat: int): void
    }
    export interface ThreeDFormat {
        setVisible(pVisible: int): void
        getVisible(): int
        setDepth(pDepth: double): void
        getDepth(): double
        getExtrusionColor(): excel.ColorFormat
        setExtrusionColorType(pExtrusionColorType: int): void
        getExtrusionColorType(): int
        setPresetMaterial(pPresetMaterial: int): void
        getPresetMaterial(): int
        incrementRotationX(increment: double): void
        incrementRotationY(increment: double): void
        getPresetThreeDFormat(): int
        getPresetExtrusionDirection(): int
        setPresetLightingDirection(pPresetLightingDirection: int): void
        getPresetLightingDirection(): int
        setPresetLightingSoftness(pPresetLightingSoftness: int): void
        getPresetLightingSoftness(): int
        setPerspective(pPerspective: int): void
        getPerspective(): int
        setRotationX(pRotationX: double): void
        getRotationX(): double
        setRotationY(pRotationY: double): void
        getRotationY(): double
        resetRotation(): void
        setExtrusionDirection(presetExtrusionDirection: int): void
        setPresetThreeDFormat(presetThreeDFormat: int): void
    }
    export interface TickLabels {
        getName(): string
        delete(): any
        getOffset(): int
        setOffset(pOffset: int): void
        getFont(): excel.Font
        getOrientation(): int
        select(): any
        setOrientation(pOrientation: int): void
        getAlignment(): int
        setAlignment(pAlignment: int): void
        getNumberFormat(): string
        setNumberFormat(pNumberFormat: string): void
        getReadingOrder(): int
        setReadingOrder(pReadingOrder: int): void
        getDepth(): int
        getAutoScaleFont(): any
        setAutoScaleFont(pAutoScaleFont: any): void
        setNumberFormatLocal(pNumberFormatLocal: any): void
        getNumberFormatLocal(): any
        setNumberFormatLinked(pNumberFormatLinked: boolean): void
        isNumberFormatLinked(): boolean
    }
    export interface TickLabels {
        getName(): string
        delete(): any
        getOffset(): int
        setOffset(pOffset: int): void
        getFont(): excel.Font
        getOrientation(): int
        select(): any
        setOrientation(pOrientation: int): void
        getAlignment(): int
        setAlignment(pAlignment: int): void
        getNumberFormat(): string
        setNumberFormat(pNumberFormat: string): void
        getReadingOrder(): int
        setReadingOrder(pReadingOrder: int): void
        getDepth(): int
        getAutoScaleFont(): any
        setAutoScaleFont(pAutoScaleFont: any): void
        setNumberFormatLocal(pNumberFormatLocal: any): void
        getNumberFormatLocal(): any
        setNumberFormatLinked(pNumberFormatLinked: boolean): void
        isNumberFormatLinked(): boolean
    }
    export interface Top10 {
        delete(): void
        setPriority(p: int): void
        getPriority(): int
        getType(): int
        getFont(): excel.Font
        getBorders(): excel.Borders
        getNumberFormat(): any
        setNumberFormat(nf: any): void
        getPercent(): boolean
        getAppliesTo(): excel.Range
        getInterior(): excel.Interior
        setFirstPriority(): void
        setLastPriority(): void
        getStopIfTrue(): boolean
        setStopIfTrue(s: boolean): void
        modifyAppliesToRange(range: excel.Range): void
        getRank(): int
        setRank(rank: int): void
        getTopBottom(): int
        setTopBottom(type: int): void
        setPercent(percent: boolean): void
    }
    export interface Top10 {
        delete(): void
        setPriority(p: int): void
        getPriority(): int
        getType(): int
        getFont(): excel.Font
        getBorders(): excel.Borders
        getNumberFormat(): any
        setNumberFormat(nf: any): void
        getPercent(): boolean
        getAppliesTo(): excel.Range
        getInterior(): excel.Interior
        setFirstPriority(): void
        setLastPriority(): void
        getStopIfTrue(): boolean
        setStopIfTrue(s: boolean): void
        modifyAppliesToRange(range: excel.Range): void
        getRank(): int
        setRank(rank: int): void
        getTopBottom(): int
        setTopBottom(type: int): void
        setPercent(percent: boolean): void
    }
    export interface TreeviewControl {
        getHidden(): any
        setHidden(pHidden: any): void
        getDrilled(): any
        setDrilled(pDrilled: any): void
    }
    export interface TreeviewControl {
        getHidden(): any
        setHidden(pHidden: any): void
        getDrilled(): any
        setDrilled(pDrilled: any): void
    }
    export interface Trendline {
        getName(): string
        delete(): any
        setName(pName: string): void
        getType(): int
        getBorder(): excel.Border
        getIndex(): int
        setType(pType: int): void
        select(): any
        clearFormats(): any
        setForward(pForward: int): void
        isDisplayEquation(): boolean
        setDisplayEquation(pDisplayEquation: boolean): void
        isDisplayRSquared(): boolean
        setDisplayRSquared(pDisplayRSquared: boolean): void
        isInterceptIsAuto(): boolean
        setInterceptIsAuto(pInterceptIsAuto: boolean): void
        getOrder(): int
        setOrder(pOrder: int): void
        getBackward(): int
        setBackward(pBackward: int): void
        getForward(): int
        getIntercept(): double
        setIntercept(pIntercept: double): void
        isNameIsAuto(): boolean
        setNameIsAuto(pNameIsAuto: boolean): void
        getPeriod(): int
        setPeriod(pPeriod: int): void
        getDataLabel(): excel.DataLabel
    }
    export interface Trendline {
        getName(): string
        delete(): any
        setName(pName: string): void
        getType(): int
        getBorder(): excel.Border
        getIndex(): int
        setType(pType: int): void
        select(): any
        clearFormats(): any
        setForward(pForward: int): void
        isDisplayEquation(): boolean
        setDisplayEquation(pDisplayEquation: boolean): void
        isDisplayRSquared(): boolean
        setDisplayRSquared(pDisplayRSquared: boolean): void
        isInterceptIsAuto(): boolean
        setInterceptIsAuto(pInterceptIsAuto: boolean): void
        getOrder(): int
        setOrder(pOrder: int): void
        getBackward(): int
        setBackward(pBackward: int): void
        getForward(): int
        getIntercept(): double
        setIntercept(pIntercept: double): void
        isNameIsAuto(): boolean
        setNameIsAuto(pNameIsAuto: boolean): void
        getPeriod(): int
        setPeriod(pPeriod: int): void
        getDataLabel(): excel.DataLabel
    }
    export interface Trendlines {
        add(type: int, order: any, period: any, forward: any, backward: any, intercept: any, displayEquation: any, displayRSquared: any, name: any): excel.Trendline
        item(index: any): excel.Trendline
        getCount(): int
    }
    export interface Trendlines {
        add(type: int, order: any, period: any, forward: any, backward: any, intercept: any, displayEquation: any, displayRSquared: any, name: any): excel.Trendline
        item(index: any): excel.Trendline
        getCount(): int
    }
    export interface UniqueValues {
        getDupeUnique(): int
        setDupeUnique(dupeUniqueType: int): void
    }
    export interface UniqueValues {
        getDupeUnique(): int
        setDupeUnique(dupeUniqueType: int): void
    }
    export interface UpBars {
        getName(): string
        delete(): any
        getBorder(): excel.Border
        select(): any
        getFill(): excel.ChartFillFormat
        getInterior(): excel.Interior
    }
    export interface UpBars {
        getName(): string
        delete(): any
        getBorder(): excel.Border
        select(): any
        getFill(): excel.ChartFillFormat
        getInterior(): excel.Interior
    }
    export interface UsedObjects {
        item(obj: any): any
        getCount(): int
        getObjects(): any
    }
    export interface UsedObjects {
        item(obj: any): any
        getCount(): int
        getObjects(): any
    }
    export interface UserAccess {
        getName(): string
        delete(): void
        setAllowEdit(pAllowEdit: boolean): void
        isAllowEdit(): boolean
    }
    export interface UserAccess {
        getName(): string
        delete(): void
        setAllowEdit(pAllowEdit: boolean): void
        isAllowEdit(): boolean
    }
    export interface UserAccessList {
        add(name: string, allowEdit: boolean): excel.UserAccess
        item(index: any): excel.UserAccess
        getCount(): int
        deleteAll(): void
    }
    export interface UserAccessList {
        add(name: string, allowEdit: boolean): excel.UserAccess
        item(index: any): excel.UserAccess
        getCount(): int
        deleteAll(): void
    }
    export interface Validation {
        add(type: int, alertStyle: any, operator: any, formula1: any, formula2: any): void
        delete(): void
        getType(): int
        modify(type: int, alertStyle: any, operator: any, formula1: any, formula2: any): void
        setIMEMode(pIMEMode: int): void
        getIMEMode(): int
        getFormula1(): string
        getFormula2(): string
        getOperator(): int
        setInCellDropdown(pInCellDropdown: boolean): void
        isValue(): boolean
        getAlertStyle(): int
        getErrorMessage(): string
        setErrorMessage(pErrorMessage: string): void
        getErrorTitle(): string
        setErrorTitle(pErrorTitle: string): void
        isIgnoreBlank(): boolean
        setIgnoreBlank(pIgnoreBlank: boolean): void
        isInCellDropdown(): boolean
        getInputMessage(): string
        setInputMessage(pInputMessage: string): void
        getInputTitle(): string
        setInputTitle(pInputTitle: string): void
        isShowError(): boolean
        setShowError(pShowError: boolean): void
        isShowInput(): boolean
        setShowInput(pShowInput: boolean): void
    }
    export interface Validation {
        add(type: int, alertStyle: any, operator: any, formula1: any, formula2: any): void
        delete(): void
        getType(): int
        modify(type: int, alertStyle: any, operator: any, formula1: any, formula2: any): void
        setIMEMode(pIMEMode: int): void
        getIMEMode(): int
        getFormula1(): string
        getFormula2(): string
        getOperator(): int
        setInCellDropdown(pInCellDropdown: boolean): void
        isValue(): boolean
        getAlertStyle(): int
        getErrorMessage(): string
        setErrorMessage(pErrorMessage: string): void
        getErrorTitle(): string
        setErrorTitle(pErrorTitle: string): void
        isIgnoreBlank(): boolean
        setIgnoreBlank(pIgnoreBlank: boolean): void
        isInCellDropdown(): boolean
        getInputMessage(): string
        setInputMessage(pInputMessage: string): void
        getInputTitle(): string
        setInputTitle(pInputTitle: string): void
        isShowError(): boolean
        setShowError(pShowError: boolean): void
        isShowInput(): boolean
        setShowInput(pShowInput: boolean): void
    }
    export interface VPageBreak {
        getLocation(): excel.Range
        delete(): void
        getType(): int
        setType(pType: int): void
        setLocation(pLocation: excel.Range): void
        getExtent(): int
        dragOff(direction: int, regionIndex: int): void
    }
    export interface VPageBreak {
        getLocation(): excel.Range
        delete(): void
        getType(): int
        setType(pType: int): void
        setLocation(pLocation: excel.Range): void
        getExtent(): int
        dragOff(direction: int, regionIndex: int): void
    }
    export interface VPageBreaks {
        add(before: any): excel.VPageBreak
        item(index: int): excel.VPageBreak
        getCount(): int
    }
    export interface VPageBreaks {
        add(before: any): excel.VPageBreak
        item(index: int): excel.VPageBreak
        getCount(): int
    }
    export interface Walls {
        getName(): string
        getBorder(): excel.Border
        paste(): void
        select(): any
        getFill(): excel.ChartFillFormat
        clearFormats(): any
        getInterior(): excel.Interior
        getPictureType(): any
        setPictureType(pPictureType: any): void
        setPictureUnit(pPictureUnit: any): void
        getPictureUnit(): any
    }
    export interface Walls {
        getName(): string
        getBorder(): excel.Border
        paste(): void
        select(): any
        getFill(): excel.ChartFillFormat
        clearFormats(): any
        getInterior(): excel.Interior
        getPictureType(): any
        setPictureType(pPictureType: any): void
        setPictureUnit(pPictureUnit: any): void
        getPictureUnit(): any
    }
    export interface Watch {
        delete(): void
        getSource(): any
    }
    export interface Watch {
        delete(): void
        getSource(): any
    }
    export interface Watches {
        add(source: any): excel.Watch
        delete(): void
        item(index: any): excel.Watch
        getCount(): int
    }
    export interface Watches {
        add(source: any): excel.Watch
        delete(): void
        item(index: any): excel.Watch
        getCount(): int
    }
    export interface WebOptions {
        getEncoding(): int
        isOrganizeInFolder(): boolean
        setOrganizeInFolder(pOrganizeInFolder: boolean): void
        setUseLongFileNames(pUseLongFileNames: boolean): void
        isUseLongFileNames(): boolean
        getFolderSuffix(): string
        isAllowPNG(): boolean
        setAllowPNG(pAllowPNG: boolean): void
        setEncoding(pEncoding: int): void
        getPixelsPerInch(): int
        setPixelsPerInch(pPixelsPerInch: int): void
        isRelyOnCSS(): boolean
        setRelyOnCSS(pRelyOnCSS: boolean): void
        isRelyOnVML(): boolean
        setRelyOnVML(pRelyOnVML: boolean): void
        getScreenSize(): int
        setScreenSize(pScreenSize: int): void
        setTargetBrowser(pTargetBrowser: int): void
        getTargetBrowser(): int
        useDefaultFolderSuffix(): void
        setDownloadComponents(pDownloadComponents: boolean): void
        isDownloadComponents(): boolean
        getLocationOfComponents(): string
        setLocationOfComponents(pLocationOfComponents: string): void
    }
    export interface WebOptions {
        getEncoding(): int
        isOrganizeInFolder(): boolean
        setOrganizeInFolder(pOrganizeInFolder: boolean): void
        setUseLongFileNames(pUseLongFileNames: boolean): void
        isUseLongFileNames(): boolean
        getFolderSuffix(): string
        isAllowPNG(): boolean
        setAllowPNG(pAllowPNG: boolean): void
        setEncoding(pEncoding: int): void
        getPixelsPerInch(): int
        setPixelsPerInch(pPixelsPerInch: int): void
        isRelyOnCSS(): boolean
        setRelyOnCSS(pRelyOnCSS: boolean): void
        isRelyOnVML(): boolean
        setRelyOnVML(pRelyOnVML: boolean): void
        getScreenSize(): int
        setScreenSize(pScreenSize: int): void
        setTargetBrowser(pTargetBrowser: int): void
        getTargetBrowser(): int
        useDefaultFolderSuffix(): void
        setDownloadComponents(pDownloadComponents: boolean): void
        isDownloadComponents(): boolean
        getLocationOfComponents(): string
        setLocationOfComponents(pLocationOfComponents: string): void
    }
    export interface Window {
        close(saveChanges: any, filename: any, routeWorkbook: any): boolean
        getType(): int
        getIndex(): int
        getWidth(): double
        getLeft(): double
        setLeft(pLeft: double): void
        setTop(pTop: double): void
        getTop(): double
        setWidth(pWidth: double): void
        activate(): any
        setCaption(pCaption: any): void
        getCaption(): any
        getHeight(): double
        setHeight(pHeight: double): void
        printPreview(index: any): any
        getSelection(): any
        getUsableHeight(): double
        getUsableWidth(): double
        isVisible(): boolean
        setVisible(pVisible: boolean): void
        getWindowState(): int
        setWindowState(pWindowState: int): void
        newWindow(): excel.Window
        printOut(from: any, to: any, copies: any, preview: any, activePrinter: any, printToFile: any, collate: any, prToFileName: any): any
        rangeFromPoint(x: int, y: int): any
        getView(): int
        getZoom(): any
        getActivePane(): excel.Pane
        largeScroll(down: any, up: any, toRight: any, toLeft: any): any
        smallScroll(down: any, up: any, toRight: any, toLeft: any): any
        getPanes(): excel.Panes
        setSplit(pSplit: boolean): void
        isSplit(): boolean
        setView(pView: int): void
        isDisplayHorizontalScrollBar(): boolean
        isDisplayVerticalScrollBar(): boolean
        setDisplayHorizontalScrollBar(pDisplayHorizontalScrollBar: boolean): void
        setDisplayVerticalScrollBar(pDisplayVerticalScrollBar: boolean): void
        setZoom(pZoom: any): void
        getMWindow(): any
        setSplitVertical(pSplitVertical: double): void
        getSplitVertical(): double
        getWindowNumber(): int
        scrollIntoView(left: int, top: int, width: int, height: int, start: any): void
        setOnWindow(pOnWindow: string): void
        getOnWindow(): string
        getSheetViews(): excel.SheetViews
        setDisplayZeros(pDisplayZeros: boolean): void
        isDisplayZeros(): boolean
        setEnableResize(pEnableResize: boolean): void
        isEnableResize(): boolean
        setFreezePanes(pFreezePanes: boolean): void
        isFreezePanes(): boolean
        setGridlineColor(pGridlineColor: int): void
        getGridlineColor(): int
        getVisibleRange(): excel.Range
        getSplitColumn(): int
        getSplitRow(): int
        setDisplayOutline(pDisplayOutline: boolean): void
        setDisplayRightToLeft(pDisplayRightToLeft: boolean): void
        getGridlineColorIndex(): int
        setGridlineColorIndex(pGridlineColorIndex: int): void
        getRangeSelection(): excel.Range
        getSelectedSheets(): excel.Sheets
        getSplitHorizontal(): double
        setSplitHorizontal(pSplitHorizontal: double): void
        scrollWorkbookTabs(sheets: any, position: any): any
        getActiveCell(): excel.Range
        getActiveChart(): excel.Chart
        getActiveSheet(): any
        getActiveWorksheetView(): excel.WorksheetView
        setDisplayFormulas(pDisplayFormulas: boolean): void
        isDisplayFormulas(): boolean
        setDisplayGridlines(pDisplayGridlines: boolean): void
        isDisplayGridlines(): boolean
        setDisplayHeadings(pDisplayHeadings: boolean): void
        isDisplayHeadings(): boolean
        setDisplayWorkbookTabs(pDisplayWorkbookTabs: boolean): void
        isDisplayWorkbookTabs(): boolean
        pointsToScreenPixelsX(points: int): int
        pointsToScreenPixelsY(points: int): int
        isDisplayOutline(): boolean
        setSplitColumn(pSplitColumn: int): void
        setSplitRow(pSplitRow: int): void
        getTabRatio(): double
        setTabRatio(pTabRatio: double): void
        activateNext(): any
        activatePrevious(): any
        isDisplayRightToLeft(): boolean
        getScrollColumn(): int
        setScrollColumn(pScrollColumn: int): void
        getScrollRow(): int
        setScrollRow(pScrollRow: int): void
    }
    export interface Window {
        close(saveChanges: any, filename: any, routeWorkbook: any): boolean
        getType(): int
        getIndex(): int
        getWidth(): double
        getLeft(): double
        setLeft(pLeft: double): void
        setTop(pTop: double): void
        getTop(): double
        setWidth(pWidth: double): void
        activate(): any
        setCaption(pCaption: any): void
        getCaption(): any
        getHeight(): double
        setHeight(pHeight: double): void
        printPreview(index: any): any
        getSelection(): any
        getUsableHeight(): double
        getUsableWidth(): double
        isVisible(): boolean
        setVisible(pVisible: boolean): void
        getWindowState(): int
        setWindowState(pWindowState: int): void
        newWindow(): excel.Window
        printOut(from: any, to: any, copies: any, preview: any, activePrinter: any, printToFile: any, collate: any, prToFileName: any): any
        rangeFromPoint(x: int, y: int): any
        getView(): int
        getZoom(): any
        getActivePane(): excel.Pane
        largeScroll(down: any, up: any, toRight: any, toLeft: any): any
        smallScroll(down: any, up: any, toRight: any, toLeft: any): any
        getPanes(): excel.Panes
        setSplit(pSplit: boolean): void
        isSplit(): boolean
        setView(pView: int): void
        isDisplayHorizontalScrollBar(): boolean
        isDisplayVerticalScrollBar(): boolean
        setDisplayHorizontalScrollBar(pDisplayHorizontalScrollBar: boolean): void
        setDisplayVerticalScrollBar(pDisplayVerticalScrollBar: boolean): void
        setZoom(pZoom: any): void
        getMWindow(): any
        setSplitVertical(pSplitVertical: double): void
        getSplitVertical(): double
        getWindowNumber(): int
        scrollIntoView(left: int, top: int, width: int, height: int, start: any): void
        setOnWindow(pOnWindow: string): void
        getOnWindow(): string
        getSheetViews(): excel.SheetViews
        setDisplayZeros(pDisplayZeros: boolean): void
        isDisplayZeros(): boolean
        setEnableResize(pEnableResize: boolean): void
        isEnableResize(): boolean
        setFreezePanes(pFreezePanes: boolean): void
        isFreezePanes(): boolean
        setGridlineColor(pGridlineColor: int): void
        getGridlineColor(): int
        getVisibleRange(): excel.Range
        getSplitColumn(): int
        getSplitRow(): int
        setDisplayOutline(pDisplayOutline: boolean): void
        setDisplayRightToLeft(pDisplayRightToLeft: boolean): void
        getGridlineColorIndex(): int
        setGridlineColorIndex(pGridlineColorIndex: int): void
        getRangeSelection(): excel.Range
        getSelectedSheets(): excel.Sheets
        getSplitHorizontal(): double
        setSplitHorizontal(pSplitHorizontal: double): void
        scrollWorkbookTabs(sheets: any, position: any): any
        getActiveCell(): excel.Range
        getActiveChart(): excel.Chart
        getActiveSheet(): any
        getActiveWorksheetView(): excel.WorksheetView
        setDisplayFormulas(pDisplayFormulas: boolean): void
        isDisplayFormulas(): boolean
        setDisplayGridlines(pDisplayGridlines: boolean): void
        isDisplayGridlines(): boolean
        setDisplayHeadings(pDisplayHeadings: boolean): void
        isDisplayHeadings(): boolean
        setDisplayWorkbookTabs(pDisplayWorkbookTabs: boolean): void
        isDisplayWorkbookTabs(): boolean
        pointsToScreenPixelsX(points: int): int
        pointsToScreenPixelsY(points: int): int
        isDisplayOutline(): boolean
        setSplitColumn(pSplitColumn: int): void
        setSplitRow(pSplitRow: int): void
        getTabRatio(): double
        setTabRatio(pTabRatio: double): void
        activateNext(): any
        activatePrevious(): any
        isDisplayRightToLeft(): boolean
        getScrollColumn(): int
        setScrollColumn(pScrollColumn: int): void
        getScrollRow(): int
        setScrollRow(pScrollRow: int): void
    }
    export interface Windows {
        item(index: any): excel.Window
        createWindow(mWindow: any): excel.Window
        getCount(): int
        getActiveWindow(): excel.Window
        breakSideBySide(): boolean
        arrange(arrangeStyle: int, activeWorkbook: any, syncHorizontal: any, syncVertical: any): any
        compareSideBySideWith(): boolean
        compareSideBySideWith(name: any): boolean
        resetPositionsSideBySide(): void
        setSyncScrollingSideBySide(pSyncScrollingSideBySide: boolean): void
        isSyncScrollingSideBySide(): boolean
    }
    export interface Windows {
        item(index: any): excel.Window
        createWindow(mWindow: any): excel.Window
        getCount(): int
        getActiveWindow(): excel.Window
        breakSideBySide(): boolean
        arrange(arrangeStyle: int, activeWorkbook: any, syncHorizontal: any, syncVertical: any): any
        compareSideBySideWith(): boolean
        compareSideBySideWith(name: any): boolean
        resetPositionsSideBySide(): void
        setSyncScrollingSideBySide(pSyncScrollingSideBySide: boolean): void
        isSyncScrollingSideBySide(): boolean
    }
    export interface Workbook {
        getName(): string
        save(): void
        close(saveChanges: any, filename: any, routeWorkbook: any): void
        getPath(): string
        getPermission(): office.Permission
        isReadOnly(): boolean
        getContainer(): any
        activate(): void
        getCommandBars(): office.CommandBars
        getWindows(): excel.Windows
        printPreview(EnableChanges: any): void
        newWindow(): excel.Window
        printOut(from: any, to: any, copies: any, preview: any, activePrinter: any, printToFile: any, collate: any, prToFileName: any): void
        printOut(from: any, to: any, copies: any, preview: any, activePrinter: any, printToFile: any, collate: any, prToFileName: any, ignorePrintAreas: any): void
        removeAllListeners(): void
        getNativeHandle(): int
        setNativeHandle(handle: int): void
        saveAs(fileName: any): void
        saveAs(fileName: any, fileFormat: any, password: any, writeResPassword: any, readOnlyRecommended: any, createBackup: any, accessMode: int, conflictResolution: any, addToMru: any, textCodepage: any, textVisualLayout: any, local: any): void
        getFullName(): string
        getMWorkbook(): any
        rejectAllChanges(when: any, who: any, where: any): void
        access$0(arg0: excel.Workbook): excel.Workbooks
        isCreateBackup(): boolean
        setHasRoutingSlip(pHasRoutingSlip: boolean): void
        setReadOnlyRecommended(pReadOnlyRecommended: boolean): void
        isReadOnlyRecommended(): boolean
        sendFaxOverInternet(recipients: any, subject: any, showMessage: any): void
        breakLink(name: string, type: int): void
        getMailer(): void
        recheckSmartTags(): void
        replyWithChanges(showMessage: any): void
        getRoutingSlip(): excel.RoutingSlip
        setPassword(pPassword: string): void
        setWritePassword(pWritePassword: string): void
        sendForReview(recipients: any, subject: any, showMessage: any, includeAttachment: any): void
        getSmartDocument(): office.SmartDocument
        getStyles(): excel.Styles
        unprotect(password: any): void
        isVBASigned(): boolean
        getWebOptions(): excel.WebOptions
        webPagePreview(): void
        isWriteReserved(): boolean
        getVBProject(): vbide.VBProject
        getWorksheets(): excel.Sheets
        getCharts(): excel.Sheets
        isRemovePersonalInformation(): boolean
        setRemovePersonalInformation(pRemovePersonalInformation: boolean): void
        saveAs2(fileName: any, fileFormat: any, password: any, writeResPassword: any, readOnlyRecommended: any, createBackup: any, accessMode: int, conflictResolution: int, addToMru: any, textCodepage: any, textVisualLayout: any, local: any): void
        saveAs2(fileName: any, fileFormat: any): void
        isSaved(): boolean
        setSaved(pSaved: boolean): void
        sendMail(recipients: any, subject: any, returnReceipt: any): void
        getSync(): office.Sync
        getSharedWorkspace(): office.SharedWorkspace
        toggleFormsDesign(): void
        exportAsFixedFormat(type: int, filename: any, quality: any, includeDocProperties: any, ignorePrintAreas: any, from: any, to: any, openAfterPublish: any, fixedFormatExtClassPtr: any): void
        isPasswordEncryptionFileProperties(): boolean
        post(destName: any): void
        protect(password: any, structure: any, windows: any): void
        reloadAs(encoding: int): void
        reply(): void
        replyAll(): void
        route(): void
        isRouted(): boolean
        addToFavorites(): void
        getCodeName(): string
        endReview(): void
        followHyperlink(address: string, subAddress: any, newWindow: any, addHistory: any, extraInfo: any, method: any, headerInfo: any): void
        forwardMailer(): void
        hasPassword(): boolean
        hasRoutingSlip(): boolean
        getHTMLProject(): excel.HTMLProject
        getCustomDocumentProperties(): any
        getDocumentLibraryVersions(): office.DocumentLibraryVersions
        getPasswordEncryptionAlgorithm(): string
        getPasswordEncryptionKeyLength(): int
        getPasswordEncryptionProvider(): string
        isEnvelopeVisible(): boolean
        setEnvelopeVisible(pEnvelopeVisible: boolean): void
        getPassword(): string
        setAcceptLabelsInFormulas(pAcceptLabelsInFormulas: boolean): void
        isAcceptLabelsInFormulas(): boolean
        getAutoUpdateFrequency(): int
        setAutoUpdateFrequency(pAutoUpdateFrequency: int): void
        isAutoUpdateSaveChanges(): boolean
        setAutoUpdateSaveChanges(pAutoUpdateSaveChanges: boolean): void
        getChangeHistoryDuration(): int
        setChangeHistoryDuration(pChangeHistoryDuration: int): void
        getConflictResolution(): int
        setConflictResolution(pConflictResolution: int): void
        deleteNumberFormat(numberFormat: string): void
        getDisplayDrawingObjects(): int
        setDisplayDrawingObjects(pDisplayDrawingObjects: int): void
        isDisplayInkComments(): boolean
        setDisplayInkComments(pDisplayInkComments: boolean): void
        isEnableAutoRecover(): boolean
        setEnableAutoRecover(pEnableAutoRecover: boolean): void
        setForceFullCalculation(forceFullCalculation: boolean): void
        isForceFullCalculation(): boolean
        getFullNameURLEncoded(): string
        highlightChangesOptions(when: any, who: any, where: any): void
        isKeepChangeHistory(): boolean
        setKeepChangeHistory(pKeepChangeHistory: boolean): void
        isListChangesOnNewSheet(): boolean
        setListChangesOnNewSheet(pListChangesOnNewSheet: boolean): void
        isMultiUserEditing(): boolean
        isPrecisionAsDisplayed(): boolean
        setPrecisionAsDisplayed(pPrecisionAsDisplayed: boolean): void
        isProtectStructure(): boolean
        getPublishObjects(): excel.PublishObjects
        purgeChangeHistoryNow(days: int, sharingPassword: any): void
        getRevisionNumber(): int
        setSaveLinkValues(pSaveLinkValues: boolean): void
        isShowConflictHistory(): boolean
        setShowConflictHistory(pShowConflictHistory: boolean): void
        getSmartTagOptions(): excel.SmartTagOptions
        isTemplateRemoveExtData(): boolean
        setTemplateRemoveExtData(pTemplateRemoveExtData: boolean): void
        isUpdateRemoteReferences(): boolean
        getWriteReservedBy(): string
        getNames(): excel.Names
        getActiveChart(): excel.Chart
        getActiveSheet(): any
        getSheets(): excel.Sheets
        getSheets(isWorksheets: boolean): excel.Sheets
        getCalculationVersion(): int
        getExcel4IntlMacroSheets(): excel.Sheets
        getExcel4MacroSheets(): excel.Sheets
        addWorkbookListener(l: excel.event.WorkbookListener): void
        removeWorkbookListener(l: excel.event.WorkbookListener): void
        saveCopyAs(filename: any): void
        checkIn(saveChanges: any, comments: any, makePublic: any): void
        isAddin(): boolean
        setAddin(pAddin: boolean): void
        linkInfo(name: string, linkInfo: int, type: any, editionRef: any): void
        AcceptAllChanges(when: any, who: any, where: any): void
        canCheckIn(): boolean
        changeFileAccess(mode: int, writePassword: any, notify: any): void
        changeLink(name: string, newName: string, type: int): void
        getColors(index: any): any
        setColors(pColors: any, index: any): void
        getCustomViews(): excel.CustomViews
        isDate1904(): boolean
        setDate1904(pDate1904: boolean): void
        exclusiveAccess(): boolean
        getFileFormat(): int
        isInplace(): boolean
        linkSources(type: any): any
        mergeWorkbook(filename: any): void
        openLinks(name: string, readOnly: any, type: any): void
        pivotCaches(): excel.PivotCaches
        protectSharing(filename: any, password: any, writeResPassword: any, readOnlyRecommended: any, createBackup: any, sharingPassword: any, fileFormat: any): void
        isProtectWindows(): boolean
        refreshAll(): void
        removeUser(index: int): void
        resetColors(): void
        runAutoMacros(which: int): void
        saveAsXMLData(filename: string, map: excel.XmlMap): void
        isSaveLinkValues(): boolean
        sendMailer(fileFormat: any, priority: int): void
        setLinkOnData(name: string, procedure: any): void
        unprotectSharing(sharingPassword: any): void
        updateFromFile(): void
        updateLink(name: any, type: any): void
        getUpdateLinks(): int
        setUpdateLinks(pUpdateLinks: int): void
        getUserStatus(): any
        getWritePassword(): string
        xmlImport(url: string, importMap: excel.XmlMap, overwrite: any, destination: any): int
        xmlImportXml(data: string, importMap: excel.XmlMap, overwrite: any, destination: any): int
        getXmlMaps(): excel.XmlMaps
        getXmlNamespaces(): excel.XmlNamespaces
        getIconSets(): excel.IconSets
        exitEditMode(): void
        printOut1(from: any, to: any, copies: any, preview: any, activePrinter: any, printToFile: any, collate: any, prToFileName: any): void
        getSheetMap(): any
        getBuiltinDocumentProperties(): any
        isHighlightChangesOnScreen(): boolean
        setHighlightChangesOnScreen(pHighlightChangesOnScreen: boolean): void
        isInactiveListBorderVisible(): boolean
        setInactiveListBorderVisible(pInactiveListBorderVisible: boolean): void
        isPersonalViewListSettings(): boolean
        setPersonalViewListSettings(pPersonalViewListSettings: boolean): void
        isPersonalViewPrintSettings(): boolean
        setPersonalViewPrintSettings(pPersonalViewPrintSettings: boolean): void
        SetPasswordEncryptionOptions(passwordEncryptionProvider: any, passwordEncryptionAlgorithm: any, passwordEncryptionKeyLength: any, passwordEncryptionFileProperties: any): void
        isShowPivotTableFieldList(): boolean
        setShowPivotTableFieldList(pShowPivotTableFieldList: boolean): void
        setUpdateRemoteReferences(pUpdateRemoteReferences: boolean): void
    }
    export interface Workbook {
        getName(): string
        save(): void
        close(saveChanges: any, filename: any, routeWorkbook: any): void
        getPath(): string
        getPermission(): office.Permission
        isReadOnly(): boolean
        getContainer(): any
        activate(): void
        getCommandBars(): office.CommandBars
        getWindows(): excel.Windows
        printPreview(EnableChanges: any): void
        newWindow(): excel.Window
        printOut(from: any, to: any, copies: any, preview: any, activePrinter: any, printToFile: any, collate: any, prToFileName: any): void
        printOut(from: any, to: any, copies: any, preview: any, activePrinter: any, printToFile: any, collate: any, prToFileName: any, ignorePrintAreas: any): void
        removeAllListeners(): void
        getNativeHandle(): int
        setNativeHandle(handle: int): void
        saveAs(fileName: any): void
        saveAs(fileName: any, fileFormat: any, password: any, writeResPassword: any, readOnlyRecommended: any, createBackup: any, accessMode: int, conflictResolution: any, addToMru: any, textCodepage: any, textVisualLayout: any, local: any): void
        getFullName(): string
        getMWorkbook(): any
        rejectAllChanges(when: any, who: any, where: any): void
        access$0(arg0: excel.Workbook): excel.Workbooks
        isCreateBackup(): boolean
        setHasRoutingSlip(pHasRoutingSlip: boolean): void
        setReadOnlyRecommended(pReadOnlyRecommended: boolean): void
        isReadOnlyRecommended(): boolean
        sendFaxOverInternet(recipients: any, subject: any, showMessage: any): void
        breakLink(name: string, type: int): void
        getMailer(): void
        recheckSmartTags(): void
        replyWithChanges(showMessage: any): void
        getRoutingSlip(): excel.RoutingSlip
        setPassword(pPassword: string): void
        setWritePassword(pWritePassword: string): void
        sendForReview(recipients: any, subject: any, showMessage: any, includeAttachment: any): void
        getSmartDocument(): office.SmartDocument
        getStyles(): excel.Styles
        unprotect(password: any): void
        isVBASigned(): boolean
        getWebOptions(): excel.WebOptions
        webPagePreview(): void
        isWriteReserved(): boolean
        getVBProject(): vbide.VBProject
        getWorksheets(): excel.Sheets
        getCharts(): excel.Sheets
        isRemovePersonalInformation(): boolean
        setRemovePersonalInformation(pRemovePersonalInformation: boolean): void
        saveAs2(fileName: any, fileFormat: any, password: any, writeResPassword: any, readOnlyRecommended: any, createBackup: any, accessMode: int, conflictResolution: int, addToMru: any, textCodepage: any, textVisualLayout: any, local: any): void
        saveAs2(fileName: any, fileFormat: any): void
        isSaved(): boolean
        setSaved(pSaved: boolean): void
        sendMail(recipients: any, subject: any, returnReceipt: any): void
        getSync(): office.Sync
        getSharedWorkspace(): office.SharedWorkspace
        toggleFormsDesign(): void
        exportAsFixedFormat(type: int, filename: any, quality: any, includeDocProperties: any, ignorePrintAreas: any, from: any, to: any, openAfterPublish: any, fixedFormatExtClassPtr: any): void
        isPasswordEncryptionFileProperties(): boolean
        post(destName: any): void
        protect(password: any, structure: any, windows: any): void
        reloadAs(encoding: int): void
        reply(): void
        replyAll(): void
        route(): void
        isRouted(): boolean
        addToFavorites(): void
        getCodeName(): string
        endReview(): void
        followHyperlink(address: string, subAddress: any, newWindow: any, addHistory: any, extraInfo: any, method: any, headerInfo: any): void
        forwardMailer(): void
        hasPassword(): boolean
        hasRoutingSlip(): boolean
        getHTMLProject(): excel.HTMLProject
        getCustomDocumentProperties(): any
        getDocumentLibraryVersions(): office.DocumentLibraryVersions
        getPasswordEncryptionAlgorithm(): string
        getPasswordEncryptionKeyLength(): int
        getPasswordEncryptionProvider(): string
        isEnvelopeVisible(): boolean
        setEnvelopeVisible(pEnvelopeVisible: boolean): void
        getPassword(): string
        setAcceptLabelsInFormulas(pAcceptLabelsInFormulas: boolean): void
        isAcceptLabelsInFormulas(): boolean
        getAutoUpdateFrequency(): int
        setAutoUpdateFrequency(pAutoUpdateFrequency: int): void
        isAutoUpdateSaveChanges(): boolean
        setAutoUpdateSaveChanges(pAutoUpdateSaveChanges: boolean): void
        getChangeHistoryDuration(): int
        setChangeHistoryDuration(pChangeHistoryDuration: int): void
        getConflictResolution(): int
        setConflictResolution(pConflictResolution: int): void
        deleteNumberFormat(numberFormat: string): void
        getDisplayDrawingObjects(): int
        setDisplayDrawingObjects(pDisplayDrawingObjects: int): void
        isDisplayInkComments(): boolean
        setDisplayInkComments(pDisplayInkComments: boolean): void
        isEnableAutoRecover(): boolean
        setEnableAutoRecover(pEnableAutoRecover: boolean): void
        setForceFullCalculation(forceFullCalculation: boolean): void
        isForceFullCalculation(): boolean
        getFullNameURLEncoded(): string
        highlightChangesOptions(when: any, who: any, where: any): void
        isKeepChangeHistory(): boolean
        setKeepChangeHistory(pKeepChangeHistory: boolean): void
        isListChangesOnNewSheet(): boolean
        setListChangesOnNewSheet(pListChangesOnNewSheet: boolean): void
        isMultiUserEditing(): boolean
        isPrecisionAsDisplayed(): boolean
        setPrecisionAsDisplayed(pPrecisionAsDisplayed: boolean): void
        isProtectStructure(): boolean
        getPublishObjects(): excel.PublishObjects
        purgeChangeHistoryNow(days: int, sharingPassword: any): void
        getRevisionNumber(): int
        setSaveLinkValues(pSaveLinkValues: boolean): void
        isShowConflictHistory(): boolean
        setShowConflictHistory(pShowConflictHistory: boolean): void
        getSmartTagOptions(): excel.SmartTagOptions
        isTemplateRemoveExtData(): boolean
        setTemplateRemoveExtData(pTemplateRemoveExtData: boolean): void
        isUpdateRemoteReferences(): boolean
        getWriteReservedBy(): string
        getNames(): excel.Names
        getActiveChart(): excel.Chart
        getActiveSheet(): any
        getSheets(): excel.Sheets
        getSheets(isWorksheets: boolean): excel.Sheets
        getCalculationVersion(): int
        getExcel4IntlMacroSheets(): excel.Sheets
        getExcel4MacroSheets(): excel.Sheets
        addWorkbookListener(l: excel.event.WorkbookListener): void
        removeWorkbookListener(l: excel.event.WorkbookListener): void
        saveCopyAs(filename: any): void
        checkIn(saveChanges: any, comments: any, makePublic: any): void
        isAddin(): boolean
        setAddin(pAddin: boolean): void
        linkInfo(name: string, linkInfo: int, type: any, editionRef: any): void
        AcceptAllChanges(when: any, who: any, where: any): void
        canCheckIn(): boolean
        changeFileAccess(mode: int, writePassword: any, notify: any): void
        changeLink(name: string, newName: string, type: int): void
        getColors(index: any): any
        setColors(pColors: any, index: any): void
        getCustomViews(): excel.CustomViews
        isDate1904(): boolean
        setDate1904(pDate1904: boolean): void
        exclusiveAccess(): boolean
        getFileFormat(): int
        isInplace(): boolean
        linkSources(type: any): any
        mergeWorkbook(filename: any): void
        openLinks(name: string, readOnly: any, type: any): void
        pivotCaches(): excel.PivotCaches
        protectSharing(filename: any, password: any, writeResPassword: any, readOnlyRecommended: any, createBackup: any, sharingPassword: any, fileFormat: any): void
        isProtectWindows(): boolean
        refreshAll(): void
        removeUser(index: int): void
        resetColors(): void
        runAutoMacros(which: int): void
        saveAsXMLData(filename: string, map: excel.XmlMap): void
        isSaveLinkValues(): boolean
        sendMailer(fileFormat: any, priority: int): void
        setLinkOnData(name: string, procedure: any): void
        unprotectSharing(sharingPassword: any): void
        updateFromFile(): void
        updateLink(name: any, type: any): void
        getUpdateLinks(): int
        setUpdateLinks(pUpdateLinks: int): void
        getUserStatus(): any
        getWritePassword(): string
        xmlImport(url: string, importMap: excel.XmlMap, overwrite: any, destination: any): int
        xmlImportXml(data: string, importMap: excel.XmlMap, overwrite: any, destination: any): int
        getXmlMaps(): excel.XmlMaps
        getXmlNamespaces(): excel.XmlNamespaces
        getIconSets(): excel.IconSets
        exitEditMode(): void
        printOut1(from: any, to: any, copies: any, preview: any, activePrinter: any, printToFile: any, collate: any, prToFileName: any): void
        getSheetMap(): any
        getBuiltinDocumentProperties(): any
        isHighlightChangesOnScreen(): boolean
        setHighlightChangesOnScreen(pHighlightChangesOnScreen: boolean): void
        isInactiveListBorderVisible(): boolean
        setInactiveListBorderVisible(pInactiveListBorderVisible: boolean): void
        isPersonalViewListSettings(): boolean
        setPersonalViewListSettings(pPersonalViewListSettings: boolean): void
        isPersonalViewPrintSettings(): boolean
        setPersonalViewPrintSettings(pPersonalViewPrintSettings: boolean): void
        SetPasswordEncryptionOptions(passwordEncryptionProvider: any, passwordEncryptionAlgorithm: any, passwordEncryptionKeyLength: any, passwordEncryptionFileProperties: any): void
        isShowPivotTableFieldList(): boolean
        setShowPivotTableFieldList(pShowPivotTableFieldList: boolean): void
        setUpdateRemoteReferences(pUpdateRemoteReferences: boolean): void
    }
    export interface Workbooks {
        add(template: any): excel.Workbook
        close(): void
        close(saveChanges: any, filename: any, routeWorkbook: any): void
        open(filename: string, updateLinks: any, readOnly: any, format: any, password: any, writeResPassword: any, ignoreReadOnlyRecommended: any, origin: any, delimiter: any, editable: any, notify: any, converter: any, addToMru: any, local: any, corruptLoad: any): excel.Workbook
        item(index: any): excel.Workbook
        getCount(): int
        getItem(index: any): excel.Workbook
        checkOut(filename: string): void
        canCheckOut(filename: string): boolean
        getWorkbook(mbook: any): excel.Workbook
        getMWorkbooks(): any
        openText(fileName: string, origin: any, startRow: any, dataType: any, textQualifier: int, consecutiveDelimiter: any, tab: any, semicolon: any, comma: any, space: any, other: any, otherChar: any, fieldInfo: any, textVisualLayout: any, decimalSeparator: any, thousandsSeparator: any, trailingMinusNumbers: any, local: any): void
        openXML(fileName: string, stylesheets: any, loadOption: any): excel.Workbook
        createWorkbook(mbook: any): excel.Workbook
        openDatabase(fileName: string, commandText: any, commandType: any, backgroundQuery: any, importDataAs: any): excel.Workbook
    }
    export interface Workbooks {
        add(template: any): excel.Workbook
        close(): void
        close(saveChanges: any, filename: any, routeWorkbook: any): void
        open(filename: string, updateLinks: any, readOnly: any, format: any, password: any, writeResPassword: any, ignoreReadOnlyRecommended: any, origin: any, delimiter: any, editable: any, notify: any, converter: any, addToMru: any, local: any, corruptLoad: any): excel.Workbook
        item(index: any): excel.Workbook
        getCount(): int
        getItem(index: any): excel.Workbook
        checkOut(filename: string): void
        canCheckOut(filename: string): boolean
        getWorkbook(mbook: any): excel.Workbook
        getMWorkbooks(): any
        openText(fileName: string, origin: any, startRow: any, dataType: any, textQualifier: int, consecutiveDelimiter: any, tab: any, semicolon: any, comma: any, space: any, other: any, otherChar: any, fieldInfo: any, textVisualLayout: any, decimalSeparator: any, thousandsSeparator: any, trailingMinusNumbers: any, local: any): void
        openXML(fileName: string, stylesheets: any, loadOption: any): excel.Workbook
        createWorkbook(mbook: any): excel.Workbook
        openDatabase(fileName: string, commandText: any, commandType: any, backgroundQuery: any, importDataAs: any): excel.Workbook
    }
    export interface Worksheet {
        getName(): string
        delete(): void
        setName(pName: string): void
        copy(before: any, after: any): void
        getType(): int
        getIndex(): int
        activate(): void
        getMainControl(): any
        printPreview(EnableChanges: any): void
        setVisible(pVisible: int): void
        checkSpelling(customDictionary: any, ignoreUppercase: any, alwaysSuggest: any, spellLang: any): void
        move(before: any, after: any): void
        printOut(from: any, to: any, copies: any, preview: any, activePrinter: any, printToFile: any, collate: any, prToFileName: any): void
        printOut(from: any, to: any, copies: any, preview: any, activePrinter: any, printToFile: any, collate: any, prToFileName: any, ignorePrintAreas: any): void
        getNativeHandle(): int
        setNativeHandle(handle: int): void
        saveAs(fileName: string, fileFormat: any, password: any, writeResPassword: any, readOnlyRecommended: any, createBackup: any, addToMru: any, textCodepage: any, textVisualLayout: any, local: any): void
        paste(destination: any, link: any): void
        createPreviewPicForOle(): string
        requestFocus(): void
        getPrevious(): any
        access$0(arg0: excel.Worksheet): any
        access$1(arg0: excel.Worksheet, arg1: excel.event.WorksheetListener, arg2: int): boolean
        getRange(cell1: any, cell2: any): excel.Range
        select(replace: any): void
        getPageSetup(): excel.PageSetup
        getNext(): any
        getCells(): excel.Range
        getCells(row: any, column: any): excel.Range
        pasteSpecial(format: any, link: any, displayAsIcon: any, iconFileName: any, iconIndex: any, iconLabel: any, noHTMLFormatting: any): void
        getVisible(): int
        getOutline(): excel.Outline
        getColumns(): excel.Range
        getRows(): excel.Range
        getMailEnvelope(): int
        getScripts(): office.Scripts
        getShapes(): excel.Shapes
        getSmartTags(): excel.SmartTags
        unprotect(password: any): void
        exportAsFixedFormat(type: int, filename: any, quality: any, includeDocProperties: any, ignorePrintAreas: any, from: any, to: any, openAfterPublish: any, fixedFormatExtClassPtr: any): void
        protect(password: any, drawingObjects: any, contents: any, scenarios: any, userInterfaceOnly: any, allowFormattingCells: any, allowFormattingColumns: any, allowFormattingRows: any, allowInsertingColumns: any, allowInsertingRows: any, allowInsertingHyperlinks: any, allowDeletingColumns: any, allowDeletingRows: any, allowSorting: any, allowFiltering: any, allowUsingPivotTables: any): void
        getCodeName(): string
        getComments(): excel.Comments
        getHyperlinks(): excel.Hyperlinks
        calculate(): void
        setDisplayRightToLeft(pDisplayRightToLeft: boolean): void
        getCircularReference(): excel.Range
        getConsolidationFunction(): int
        getConsolidationOptions(): any
        getConsolidationSources(): any
        getCustomProperties(): excel.CustomProperties
        isDisplayPageBreaks(): boolean
        isEnableAutoFilter(): boolean
        isEnableCalculation(): boolean
        isEnableOutlining(): boolean
        isEnablePivotTable(): boolean
        getEnableSelection(): int
        isProtectScenarios(): boolean
        getStandardHeight(): double
        isTransitionExpEval(): boolean
        isTransitionFormEntry(): boolean
        setAutoFilterMode(pAutoFilterMode: boolean): void
        setDisplayPageBreaks(pDisplayPageBreaks: boolean): void
        setEnableAutoFilter(pEnableAutoFilter: boolean): void
        setEnableCalculation(pEnableCalculation: boolean): void
        setEnableOutlining(pEnableOutlining: boolean): void
        setEnablePivotTable(pEnablePivotTable: boolean): void
        setEnableSelection(pEnableSelection: int): void
        getNames(): excel.Names
        evaluate(name: any): any
        getTab(): excel.Tab
        getMWorksheet(): any
        addWorksheetListener(l: excel.event.WorksheetListener): void
        removeWorksheetListener(l: excel.event.WorksheetListener): void
        OLEObjects(index: any): any
        scenarios(index: any): any
        getAutoFilter(): excel.AutoFilter
        getSort(): excel.Sort
        chartObjects(index: any): any
        isAutoFilterMode(): boolean
        isFilterMode(): boolean
        getListObjects(): excel.ListObjects
        getProtection(): excel.Protection
        getQueryTables(): excel.QueryTables
        getScrollArea(): string
        getStandardWidth(): double
        setScrollArea(pScrollArea: string): void
        setStandardWidth(pStandardWidth: double): void
        circleInvalid(): void
        clearArrows(): void
        clearCircles(): void
        pivotTables(index: any): any
        showDataForm(): void
        xmlDataQuery(xPath: string, selectionNamespaces: any, map: any): excel.Range
        xmlMapQuery(xPath: string, selectionNamespaces: any, map: any): excel.Range
        setObjectInfo(ObjInfo: int): void
        getObjectInfo(): int
        getHPageBreaks(): excel.HPageBreaks
        getVPageBreaks(): excel.VPageBreaks
        showAllData(): void
        isProtectionMode(): boolean
        getChartObjects(): excel.ChartObjects
        isProtectContents(): boolean
        isProtectDrawingObjects(): boolean
        setBackgroundPicture(filename: string): void
        isDisplayRightToLeft(): boolean
        setTransitionExpEval(pTransitionExpEval: boolean): void
        setTransitionFormEntry(pTransitionFormEntry: boolean): void
        getPrintedCommentPages(): int
        resetAllPageBreaks(): void
        fireWorkbookEvent(l: excel.event.WorksheetListener, event: int): boolean
        pivotTableWizard(sourceType: any, sourceData: any, tableDestination: any, tableName: any, rowGrand: any, columnGrand: any, saveData: any, hasAutoFormat: any, autoPage: any, reserved: any, backgroundQuery: any, optimizeCache: any, pageFieldOrder: any, pageFieldWrapCount: any, readData: any, connection: any): excel.PivotTable
        getUsedRange(): excel.Range
    }
    export interface Worksheet {
        getName(): string
        delete(): void
        setName(pName: string): void
        copy(before: any, after: any): void
        getType(): int
        getIndex(): int
        activate(): void
        getMainControl(): any
        printPreview(EnableChanges: any): void
        setVisible(pVisible: int): void
        checkSpelling(customDictionary: any, ignoreUppercase: any, alwaysSuggest: any, spellLang: any): void
        move(before: any, after: any): void
        printOut(from: any, to: any, copies: any, preview: any, activePrinter: any, printToFile: any, collate: any, prToFileName: any): void
        printOut(from: any, to: any, copies: any, preview: any, activePrinter: any, printToFile: any, collate: any, prToFileName: any, ignorePrintAreas: any): void
        getNativeHandle(): int
        setNativeHandle(handle: int): void
        saveAs(fileName: string, fileFormat: any, password: any, writeResPassword: any, readOnlyRecommended: any, createBackup: any, addToMru: any, textCodepage: any, textVisualLayout: any, local: any): void
        paste(destination: any, link: any): void
        createPreviewPicForOle(): string
        requestFocus(): void
        getPrevious(): any
        access$0(arg0: excel.Worksheet): any
        access$1(arg0: excel.Worksheet, arg1: excel.event.WorksheetListener, arg2: int): boolean
        getRange(cell1: any, cell2: any): excel.Range
        select(replace: any): void
        getPageSetup(): excel.PageSetup
        getNext(): any
        getCells(): excel.Range
        getCells(row: any, column: any): excel.Range
        pasteSpecial(format: any, link: any, displayAsIcon: any, iconFileName: any, iconIndex: any, iconLabel: any, noHTMLFormatting: any): void
        getVisible(): int
        getOutline(): excel.Outline
        getColumns(): excel.Range
        getRows(): excel.Range
        getMailEnvelope(): int
        getScripts(): office.Scripts
        getShapes(): excel.Shapes
        getSmartTags(): excel.SmartTags
        unprotect(password: any): void
        exportAsFixedFormat(type: int, filename: any, quality: any, includeDocProperties: any, ignorePrintAreas: any, from: any, to: any, openAfterPublish: any, fixedFormatExtClassPtr: any): void
        protect(password: any, drawingObjects: any, contents: any, scenarios: any, userInterfaceOnly: any, allowFormattingCells: any, allowFormattingColumns: any, allowFormattingRows: any, allowInsertingColumns: any, allowInsertingRows: any, allowInsertingHyperlinks: any, allowDeletingColumns: any, allowDeletingRows: any, allowSorting: any, allowFiltering: any, allowUsingPivotTables: any): void
        getCodeName(): string
        getComments(): excel.Comments
        getHyperlinks(): excel.Hyperlinks
        calculate(): void
        setDisplayRightToLeft(pDisplayRightToLeft: boolean): void
        getCircularReference(): excel.Range
        getConsolidationFunction(): int
        getConsolidationOptions(): any
        getConsolidationSources(): any
        getCustomProperties(): excel.CustomProperties
        isDisplayPageBreaks(): boolean
        isEnableAutoFilter(): boolean
        isEnableCalculation(): boolean
        isEnableOutlining(): boolean
        isEnablePivotTable(): boolean
        getEnableSelection(): int
        isProtectScenarios(): boolean
        getStandardHeight(): double
        isTransitionExpEval(): boolean
        isTransitionFormEntry(): boolean
        setAutoFilterMode(pAutoFilterMode: boolean): void
        setDisplayPageBreaks(pDisplayPageBreaks: boolean): void
        setEnableAutoFilter(pEnableAutoFilter: boolean): void
        setEnableCalculation(pEnableCalculation: boolean): void
        setEnableOutlining(pEnableOutlining: boolean): void
        setEnablePivotTable(pEnablePivotTable: boolean): void
        setEnableSelection(pEnableSelection: int): void
        getNames(): excel.Names
        evaluate(name: any): any
        getTab(): excel.Tab
        getMWorksheet(): any
        addWorksheetListener(l: excel.event.WorksheetListener): void
        removeWorksheetListener(l: excel.event.WorksheetListener): void
        OLEObjects(index: any): any
        scenarios(index: any): any
        getAutoFilter(): excel.AutoFilter
        getSort(): excel.Sort
        chartObjects(index: any): any
        isAutoFilterMode(): boolean
        isFilterMode(): boolean
        getListObjects(): excel.ListObjects
        getProtection(): excel.Protection
        getQueryTables(): excel.QueryTables
        getScrollArea(): string
        getStandardWidth(): double
        setScrollArea(pScrollArea: string): void
        setStandardWidth(pStandardWidth: double): void
        circleInvalid(): void
        clearArrows(): void
        clearCircles(): void
        pivotTables(index: any): any
        showDataForm(): void
        xmlDataQuery(xPath: string, selectionNamespaces: any, map: any): excel.Range
        xmlMapQuery(xPath: string, selectionNamespaces: any, map: any): excel.Range
        setObjectInfo(ObjInfo: int): void
        getObjectInfo(): int
        getHPageBreaks(): excel.HPageBreaks
        getVPageBreaks(): excel.VPageBreaks
        showAllData(): void
        isProtectionMode(): boolean
        getChartObjects(): excel.ChartObjects
        isProtectContents(): boolean
        isProtectDrawingObjects(): boolean
        setBackgroundPicture(filename: string): void
        isDisplayRightToLeft(): boolean
        setTransitionExpEval(pTransitionExpEval: boolean): void
        setTransitionFormEntry(pTransitionFormEntry: boolean): void
        getPrintedCommentPages(): int
        resetAllPageBreaks(): void
        fireWorkbookEvent(l: excel.event.WorksheetListener, event: int): boolean
        pivotTableWizard(sourceType: any, sourceData: any, tableDestination: any, tableName: any, rowGrand: any, columnGrand: any, saveData: any, hasAutoFormat: any, autoPage: any, reserved: any, backgroundQuery: any, optimizeCache: any, pageFieldOrder: any, pageFieldWrapCount: any, readData: any, connection: any): excel.PivotTable
        getUsedRange(): excel.Range
    }
    export interface WorksheetFunction {
        index(arg1: any, arg2: double, arg3: any, arg4: any): any
        count(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        atan2(X_num: double, Y_num: double): double
        log(number: double, base: any): double
        log10(number: double): double
        min(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        max(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        replace(originString: string, startNumber: double, charNumbers: double, newString: string): string
        trim(string: string): string
        throwException(value: any): void
        find(findString: string, withinString: string, startNumber: any): double
        clean(string: string): string
        sum(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        search(findString: string, withinString: string, startNumber: any): double
        lookup(lookup_value: any, lookup_vector: any, result_vector: any): any
        frequency(arg1: any, arg2: any): any
        asin(number: double): double
        acos(number: double): double
        floor(number: double, significance: double): double
        round(number: double, digits: double): double
        sinh(number: double): double
        cosh(number: double): double
        tanh(number: double): double
        match(lookup_Value: any, lookup_array: any, matchType: any): double
        isError(arg1: any): boolean
        and(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): boolean
        convert(number: any, from_unit: any, to_unit: any): double
        text(v: any, formatText: string): string
        mode(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        average(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        getDefaultArgByLong(o: any, defaultValue: number): string
        getActiveBinder(): any
        getDefaultArgByInt(o: any, defaultValue: any): string
        var(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        delta(number1: any, number2: any): double
        getDefaultArgByDouble(o: any, defaultValue: any): string
        subtotal(criteria: double, array: excel.Range, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        phonetic(arg1: excel.Range): string
        trend(range1: any, range2: any, new_xs: any, const1: any): any
        f_Dist(x: double, degrees_freedom1: double, degrees_freedom2: double, cumulative: boolean): double
        f_Inv(Probability: double, degrees_freedom1: double, degrees_freedom2: double): double
        f_Test(array1: any, array2: any): double
        fact(number: double): double
        fDist(x: double, degrees_freedom1: double, degrees_freedom2: double): double
        findB(findString: string, withinString: string, startNumber: any): double
        fInv(Probability: double, degrees_freedom1: double, degrees_freedom2: double): double
        fisher(x: double): double
        fixed(currencyNumber: double, decimalNumber: any, noComma: any): string
        forecast(x: double, known_ys: any, known_xs: any): double
        fTest(array1: any, array2: any): double
        pmt(rate: double, nper: double, pv: double, fv: any, type: any): double
        gammaInv(probability: double, alpha: double, beta: double): double
        gammaLn(x: double): double
        gcd(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        geoMean(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        geStep(number: any, upper_limit: any): double
        growth(known_ys: any, known_xs: any, new_xs: any, const1: any): any
        harMean(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        hex2Bin(number: any, places: any): string
        hex2Dec(number: any): string
        hex2Oct(number: any, places: any): string
        hLookup(lookup_value: any, table_array: any, row_index_num: any, range_lookup: any): any
        imAbs(inumber: any): string
        imCos(inumber: any): string
        imDiv(inumber1: any, inumber2: any): string
        imExp(inumber: any): string
        imLn(inumber: any): string
        imLog10(inumber: any): string
        imLog2(inumber: any): string
        imPower(inumber1: any, inumber2: any): string
        imReal(inumber: any): double
        imSin(inumber: any): string
        imSqrt(inumber: any): string
        imSub(inumber1: any, inumber2: any): string
        imSum(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): string
        intRate(settlement: any, maturity: any, investment: any, redemption: any, basis: any): double
        ipmt(rate: double, per: double, nper: double, pv: double, fv: any, type: any): double
        irr(value: any, guess: any): double
        isErr(arg1: any): boolean
        isEven(number: any): boolean
        isNA(arg1: any): boolean
        isNumber(arg1: any): boolean
        isOdd(number: any): boolean
        ispmt(rate: double, per: double, nper: double, pv: double): double
        isText(arg1: any): boolean
        kurt(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        large(array: any, k: double): double
        lcm(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        linEst(known_ys: any, known_xs: any, consts: any, stats: any): any
        logEst(known_ys: any, known_xs: any, consts: any, stats: any): any
        logInv(probability: double, mean: double, standard_dev: double): double
        mDeterm(array: any): double
        median(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        mInverse(array: any): any
        mIrr(value: any, frate: double, rrate: double): double
        mMult(array1: any, array2: any): any
        mRound(number: any, multiple: any): double
        nominal(effect_rate: any, npery: any): double
        norm_Inv(probability: double, mean: double, standard_dev: double): double
        normDist(x: double, mean: double, standard_dev: double, cumulative: boolean): double
        normInv(probability: double, mean: double, standard_dev: double): double
        normSInv(probability: double): double
        nPer(rate: double, pmt: double, pv: double, fv: any, type: any): double
        npv(rate: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        oct2Bin(number: any, places: any): string
        oct2Dec(number: any): string
        oct2Hex(number: any, places: any): string
        odd(number: double): double
        pearson(arg1: any, arg2: any): double
        permut(number: double, number_chosen: double): double
        poisson(x: double, mean: double, cumulative: boolean): double
        power(number: double, power: double): double
        ppmt(rate: double, per: double, nper: double, pv: double, fv: any, type: any): double
        price(settlement: any, maturity: any, rate: any, yld: any, redemption: any, frequency: any, basis: any): double
        priceMat(settlement: any, maturity: any, issue: any, rate: any, yld: any, basis: any): double
        prob(range1: any, range2: any, lower_limit: double, upper_limit: any): double
        product(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        proper(string: string): string
        quartile(range: any, quart: double): double
        quotient(numerator: any, denominator: any): double
        radians(angle: double): double
        rank_Eq(number: double, range: excel.Range, order: any): double
        received(settlement: any, maturity: any, investment: any, discount: any, basis: any): double
        replaceB(originString: string, startNumber: double, charNumbers: double, newString: string): string
        rept(text: string, number_times: double): string
        roman(number: double, form: any): string
        roundUp(number: double, digits: double): double
        rSq(range1: any, range2: any): double
        rtd(progID: any, server: any, topic1: any, topic2: any, topic3: any, topic4: any, topic5: any, topic6: any, topic7: any, topic8: any, topic9: any, topic10: any, topic11: any, topic12: any, topic13: any, topic14: any, topic15: any, topic16: any, topic17: any, topic18: any, topic19: any, topic20: any, topic21: any, topic22: any, topic23: any, topic24: any, topic25: any, topic26: any, topic27: any, topic28: any): void
        searchB(findText: string, withinText: string, startNumber: any): double
        skew(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        sln(cost: double, salvage: double, life: double): double
        slope(range1: any, range2: any): double
        small(range: any, k: double): double
        sqrtPi(number: any): double
        stDev(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        stDev_P(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        stDevP(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        stEyx(range1: any, range2: any): double
        sumIf(array: excel.Range, criteria: any, sum_array: any): double
        sumIfs(array: excel.Range, sum_array: excel.Range, criteria: any, arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any): double
        sumSq(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        sumX2MY2(array1: any, array2: any): double
        sumX2PY2(array1: any, array2: any): double
        sumXMY2(array1: any, array2: any): double
        syd(cost: double, salvage: double, life: double, per: double): double
        t_Inv_2T(probability: double, degrees_freedom: double): double
        tInv(probability: double, degrees_freedom: double): double
        t_Test(array1: any, array2: any, tails: double, type: double): double
        tBillEq(settlement: any, maturity: any, discount: any): double
        tDist(x: double, degrees_freedom: double, tails: double): double
        trimMean(range: any, percent: double): double
        tTest(array1: any, array2: any, tails: double, type: double): double
        usDollar(arg1: double, arg2: double): string
        var_P(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        varP(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        vdb(cost: double, salvage: double, life: double, start_period: double, end_period: double, factor: any, no_switch: any): double
        vLookup(lookup_value: any, table_array: any, col_index_num: any, range_lookup: any): any
        weekday(Serial_number: any, return_type: any): double
        weekNum(Serial_number: any, return_type: any): double
        weibull(x: double, alpha: double, beta: double, cumulative: boolean): double
        workday(start_date: any, days: any, holidays: any): long
        xirr(values: any, dates: any, guess: any): double
        yearFrac(start_date: any, end_date: any, basis: any): double
        yieldMat(settlement: any, maturity: any, issue: any, rate: any, pr: any, basis: any): double
        z_Test(array: any, x: double, sigma: any): double
        zTest(array: any, x: double, sigma: any): double
        rank(number: double, range: excel.Range, order: any): double
        accrInt(issue: any, first_interest: any, settlement: any, rate: any, par: any, frequency: any, basis: any): double
        rate(nper: double, pmt: double, pv: double, fv: any, type: any, guess: any): double
        accrIntM(issue: any, maturity: any, rate: any, par: any, basis: any): double
        acosh(number: double): double
        amorLinc(cost: any, date_purchased: any, first_period: any, salvage: any, period: any, rate: any, basis: any): double
        asc(string: string): string
        asinh(number: double): double
        atanh(number: double): double
        aveDev(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        bahtText(arg1: double): string
        besselI(x: any, n: any): double
        besselJ(x: any, n: any): double
        besselK(x: any, n: any): double
        besselY(x: any, n: any): double
        beta_Inv(probability: double, alpha: double, beta: double, a: any, b: any): double
        betaDist(x: double, alpha: double, beta: double, a: any, b: any): double
        betaInv(probability: double, alpha: double, beta: double, a: any, b: any): double
        bin2Dec(number: any): string
        bin2Hex(number: any, places: any): string
        bin2Oct(number: any, places: any): string
        ceiling(number: double, significance: double): double
        chiDist(x: double, degrees_freedom: double): double
        chiInv(probability: double, degrees_freedom: double): double
        chiTest(actual_range: any, expected_range: any): double
        choose(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): any
        complex(real_num: any, i_number: any, arg3: any): string
        combin(number: double, number_chosen: double): double
        correl(range1: any, range2: any): double
        countA(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        countIf(arg1: excel.Range, arg2: any): double
        countIfs(arg1: excel.Range, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        coupDays(settlement: any, maturity: any, frequency: any, basis: any): double
        coupNcd(settlement: any, maturity: any, frequency: any, basis: any): double
        coupNum(settlement: any, maturity: any, frequency: any, basis: any): double
        coupPcd(settlement: any, maturity: any, frequency: any, basis: any): double
        covar(arg1: any, arg2: any): double
        cumIPmt(rate: any, nper: any, pv: any, start_period: any, end_period: any, type: any): double
        cumPrinc(rate: any, nper: any, pv: any, start_period: any, end_period: any, type: any): double
        dAverage(database: excel.Range, field: any, criteria: any): double
        days360(start_date: any, end_date: any, method: any): double
        dbcs(arg1: string): string
        dCount(database: excel.Range, field: any, criteria: any): double
        dCountA(database: excel.Range, field: any, criteria: any): double
        ddb(cost: double, salvage: double, life: double, period: double, factor: any): double
        dec2Bin(number: any, places: any): string
        dec2Hex(number: any, places: any): string
        dec2Oct(number: any, places: any): string
        degrees(angle: double): double
        devSq(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        dGet(database: excel.Range, field: any, criteria: any): any
        disc(settlement: any, maturity: any, pr: any, redemption: any, basis: any): double
        dMax(database: excel.Range, field: any, criteria: any): double
        dMin(database: excel.Range, field: any, criteria: any): double
        dollar(currencyNumber: double, decimalNumber: any): string
        dollarDe(fractional_dollar: any, fraction: any): double
        dollarFr(decimal_dollar: any, fraction: any): double
        dProduct(database: excel.Range, field: any, criteria: any): double
        dStDev(database: excel.Range, field: any, criteria: any): double
        dStDevp(database: excel.Range, field: any, criteria: any): double
        dSum(database: excel.Range, field: any, criteria: any): double
        duration(settlement: any, maturity: any, coupon: any, yld: any, frequency: any, basis: any): double
        dVar(database: excel.Range, field: any, criteria: any): double
        dVarp(database: excel.Range, field: any, criteria: any): double
        eDate(start_date: any, method: any): double
        effect(norminal_rate: any, npery: any): double
        eoMonth(start_date: any, method: any): double
        erf(lower_limit: any, upper_limit: any): double
        erfC(x: any): double
        even(number: double): double
        intercept(arg1: any, arg2: any): double
        getRangeString(array2: any, array1: any): string
        getRangeString(array: any): string
        processFormula(formulaStr: string): any
        amorDegrc(cost: any, date_purchased: any, first_period: any, salvage: any, period: any, rate: any, basis: any): double
        processFormulas(formulaStr: string, args: any[]): any
        beta_Dist(x: double, alpha: double, beta: double, cumulative: boolean, a: any, b: any): double
        binom_Dist(number_s: double, trials: double, probability_s: double, cumulative: boolean): double
        binomDist(number_s: double, trials: double, probability_s: double, cumulative: boolean): double
        confidence(alpha: double, standard_dev: double, size: double): double
        countBlank(range: excel.Range): double
        coupDayBs(settlement: any, maturity: any, frequency: any, basis: any): double
        coupDaysNc(settlement: any, maturity: any, frequency: any, basis: any): double
        critBinom(trials: double, probability_s: double, alpha: double): double
        erf_Precise(lower_limit: any): double
        erfC_Precise(x: any): double
        expon_Dist(x: double, lambda: double, cumulative: boolean): double
        exponDist(x: double, lambda: double, cumulative: boolean): double
        f_Dist_RT(x: double, degrees_freedom1: double, degrees_freedom2: double): double
        factDouble(number: any): double
        fisherInv(y: double): double
        fVSchedule(principal: any, schedule: any): double
        gamma_Dist(x: double, alpha: double, beta: double, cumulative: boolean): double
        gamma_Inv(probability: double, alpha: double, beta: double): double
        gammaDist(x: double, alpha: double, beta: double, cumulative: boolean): double
        gammaLn_Precise(x: double): double
        hypGeomDist(sample_s: double, number_sample: double, population_s: double, number_population: double): double
        hypGeom_Dist(sample_s: double, number_sample: double, population_s: double, number_population: double, b: boolean): double
        imaginary(inumber: any): double
        imArgument(inumber: any): string
        imConjugate(inumber: any): string
        imProduct(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): string
        isLogical(arg1: any): boolean
        isNonText(arg1: any): boolean
        logNorm_Inv(probability: double, mean: double, standard_dev: double): double
        logNormDist(probability: double, mean: double, standard_dev: double): double
        getPromptString(type: int): string
        transpose(range: any): any
        mDuration(settlement: any, maturity: any, coupon: any, yld: any, frequency: any, basis: any): double
        mode_Sngl(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        multiNomial(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        negBinomDist(number_f: double, number_s: double, probability_s: double): double
        networkdays(start_date: any, end_date: any, holidays: any): double
        norm_Dist(x: double, mean: double, standard_dev: double, cumulative: boolean): double
        normSDist(z: double): double
        oddFPrice(settlement: any, maturity: any, issue: any, first_coupon: any, rate: any, yld: any, redemption: any, frequency: any, basis: any): double
        oddFYield(settlement: any, maturity: any, issue: any, first_coupon: any, rate: any, pr: any, redemption: any, frequency: any, basis: any): double
        oddLPrice(settlement: any, maturity: any, last_interest: any, rate: any, yld: any, redemption: any, frequency: any, basis: any): double
        oddLYield(settlement: any, maturity: any, last_interest: any, rate: any, pr: any, redemption: any, frequency: any, basis: any): double
        percentile(range: any, k: double): double
        percentile_Exc(range: any, k: double): double
        percentile_Inc(range: any, k: double): double
        percentRank(array: any, x: double, significance: any): double
        percentRank_Exc(array: any, x: double, significance: any): double
        percentRank_Inc(array: any, x: double, significance: any): double
        poisson_Dist(x: double, mean: double, cumulative: boolean): double
        priceDisc(settlement: any, maturity: any, discount: any, redemption: any, basis: any): double
        quartile_Exc(range: any, quart: double): double
        quartile_Inc(range: any, quart: double): double
        randBetween(bottom: any, top: any): double
        roundDown(number: double, digits: double): double
        seriesSum(x: any, n: any, m: any, coefficients: any): double
        standardize(x: double, mean: double, standard_dev: double): double
        substitute(originString: string, oldString: string, newString: string, orderNumber: any): string
        sumProduct(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        t_Dist_2T(x: double, degrees_freedom: double): double
        t_Dist_RT(x: double, degrees_freedom: double): double
        tBillPrice(settlement: any, maturity: any, discount: any): double
        tBillYield(settlement: any, maturity: any, pr: any): double
        weibull_Dist(x: double, alpha: double, beta: double, cumulative: boolean): double
        yieldDisc(settlement: any, maturity: any, pr: any, redemption: any, basis: any): double
        getArrayString(values: string[], needBrackets: boolean): string
        getArrayString(values: any): string
        getArrayString(values: number[], needBrackets: boolean): string
        getArrayString(values: number[]): string
        getArrayString(values: number[], needBrackets: boolean): string
        getArrayString(values: string[]): string
        getArrayString(values: any): string
        getArrayString(values: number[]): string
        getArrayString(values: any): string
        getArrayString(values: number[], needBrackets: boolean): string
        getArrayString(values: number[]): string
        getArrayString(values: any): string
        getArrayString(values: number[], needBrackets: boolean): string
        getArrayString(values: number[]): string
        getArrayString(values: any): string
        isFUNCTION_ERROR(str: string): boolean
    }
    export interface WorksheetFunction {
        index(arg1: any, arg2: double, arg3: any, arg4: any): any
        count(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        atan2(X_num: double, Y_num: double): double
        log(number: double, base: any): double
        log10(number: double): double
        min(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        max(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        replace(originString: string, startNumber: double, charNumbers: double, newString: string): string
        trim(string: string): string
        throwException(value: any): void
        find(findString: string, withinString: string, startNumber: any): double
        clean(string: string): string
        sum(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        search(findString: string, withinString: string, startNumber: any): double
        lookup(lookup_value: any, lookup_vector: any, result_vector: any): any
        frequency(arg1: any, arg2: any): any
        asin(number: double): double
        acos(number: double): double
        floor(number: double, significance: double): double
        round(number: double, digits: double): double
        sinh(number: double): double
        cosh(number: double): double
        tanh(number: double): double
        match(lookup_Value: any, lookup_array: any, matchType: any): double
        isError(arg1: any): boolean
        and(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): boolean
        convert(number: any, from_unit: any, to_unit: any): double
        text(v: any, formatText: string): string
        mode(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        average(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        getDefaultArgByLong(o: any, defaultValue: number): string
        getActiveBinder(): any
        getDefaultArgByInt(o: any, defaultValue: any): string
        var(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        delta(number1: any, number2: any): double
        getDefaultArgByDouble(o: any, defaultValue: any): string
        subtotal(criteria: double, array: excel.Range, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        phonetic(arg1: excel.Range): string
        trend(range1: any, range2: any, new_xs: any, const1: any): any
        f_Dist(x: double, degrees_freedom1: double, degrees_freedom2: double, cumulative: boolean): double
        f_Inv(Probability: double, degrees_freedom1: double, degrees_freedom2: double): double
        f_Test(array1: any, array2: any): double
        fact(number: double): double
        fDist(x: double, degrees_freedom1: double, degrees_freedom2: double): double
        findB(findString: string, withinString: string, startNumber: any): double
        fInv(Probability: double, degrees_freedom1: double, degrees_freedom2: double): double
        fisher(x: double): double
        fixed(currencyNumber: double, decimalNumber: any, noComma: any): string
        forecast(x: double, known_ys: any, known_xs: any): double
        fTest(array1: any, array2: any): double
        pmt(rate: double, nper: double, pv: double, fv: any, type: any): double
        gammaInv(probability: double, alpha: double, beta: double): double
        gammaLn(x: double): double
        gcd(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        geoMean(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        geStep(number: any, upper_limit: any): double
        growth(known_ys: any, known_xs: any, new_xs: any, const1: any): any
        harMean(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        hex2Bin(number: any, places: any): string
        hex2Dec(number: any): string
        hex2Oct(number: any, places: any): string
        hLookup(lookup_value: any, table_array: any, row_index_num: any, range_lookup: any): any
        imAbs(inumber: any): string
        imCos(inumber: any): string
        imDiv(inumber1: any, inumber2: any): string
        imExp(inumber: any): string
        imLn(inumber: any): string
        imLog10(inumber: any): string
        imLog2(inumber: any): string
        imPower(inumber1: any, inumber2: any): string
        imReal(inumber: any): double
        imSin(inumber: any): string
        imSqrt(inumber: any): string
        imSub(inumber1: any, inumber2: any): string
        imSum(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): string
        intRate(settlement: any, maturity: any, investment: any, redemption: any, basis: any): double
        ipmt(rate: double, per: double, nper: double, pv: double, fv: any, type: any): double
        irr(value: any, guess: any): double
        isErr(arg1: any): boolean
        isEven(number: any): boolean
        isNA(arg1: any): boolean
        isNumber(arg1: any): boolean
        isOdd(number: any): boolean
        ispmt(rate: double, per: double, nper: double, pv: double): double
        isText(arg1: any): boolean
        kurt(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        large(array: any, k: double): double
        lcm(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        linEst(known_ys: any, known_xs: any, consts: any, stats: any): any
        logEst(known_ys: any, known_xs: any, consts: any, stats: any): any
        logInv(probability: double, mean: double, standard_dev: double): double
        mDeterm(array: any): double
        median(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        mInverse(array: any): any
        mIrr(value: any, frate: double, rrate: double): double
        mMult(array1: any, array2: any): any
        mRound(number: any, multiple: any): double
        nominal(effect_rate: any, npery: any): double
        norm_Inv(probability: double, mean: double, standard_dev: double): double
        normDist(x: double, mean: double, standard_dev: double, cumulative: boolean): double
        normInv(probability: double, mean: double, standard_dev: double): double
        normSInv(probability: double): double
        nPer(rate: double, pmt: double, pv: double, fv: any, type: any): double
        npv(rate: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        oct2Bin(number: any, places: any): string
        oct2Dec(number: any): string
        oct2Hex(number: any, places: any): string
        odd(number: double): double
        pearson(arg1: any, arg2: any): double
        permut(number: double, number_chosen: double): double
        poisson(x: double, mean: double, cumulative: boolean): double
        power(number: double, power: double): double
        ppmt(rate: double, per: double, nper: double, pv: double, fv: any, type: any): double
        price(settlement: any, maturity: any, rate: any, yld: any, redemption: any, frequency: any, basis: any): double
        priceMat(settlement: any, maturity: any, issue: any, rate: any, yld: any, basis: any): double
        prob(range1: any, range2: any, lower_limit: double, upper_limit: any): double
        product(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        proper(string: string): string
        quartile(range: any, quart: double): double
        quotient(numerator: any, denominator: any): double
        radians(angle: double): double
        rank_Eq(number: double, range: excel.Range, order: any): double
        received(settlement: any, maturity: any, investment: any, discount: any, basis: any): double
        replaceB(originString: string, startNumber: double, charNumbers: double, newString: string): string
        rept(text: string, number_times: double): string
        roman(number: double, form: any): string
        roundUp(number: double, digits: double): double
        rSq(range1: any, range2: any): double
        rtd(progID: any, server: any, topic1: any, topic2: any, topic3: any, topic4: any, topic5: any, topic6: any, topic7: any, topic8: any, topic9: any, topic10: any, topic11: any, topic12: any, topic13: any, topic14: any, topic15: any, topic16: any, topic17: any, topic18: any, topic19: any, topic20: any, topic21: any, topic22: any, topic23: any, topic24: any, topic25: any, topic26: any, topic27: any, topic28: any): void
        searchB(findText: string, withinText: string, startNumber: any): double
        skew(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        sln(cost: double, salvage: double, life: double): double
        slope(range1: any, range2: any): double
        small(range: any, k: double): double
        sqrtPi(number: any): double
        stDev(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        stDev_P(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        stDevP(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        stEyx(range1: any, range2: any): double
        sumIf(array: excel.Range, criteria: any, sum_array: any): double
        sumIfs(array: excel.Range, sum_array: excel.Range, criteria: any, arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any): double
        sumSq(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        sumX2MY2(array1: any, array2: any): double
        sumX2PY2(array1: any, array2: any): double
        sumXMY2(array1: any, array2: any): double
        syd(cost: double, salvage: double, life: double, per: double): double
        t_Inv_2T(probability: double, degrees_freedom: double): double
        tInv(probability: double, degrees_freedom: double): double
        t_Test(array1: any, array2: any, tails: double, type: double): double
        tBillEq(settlement: any, maturity: any, discount: any): double
        tDist(x: double, degrees_freedom: double, tails: double): double
        trimMean(range: any, percent: double): double
        tTest(array1: any, array2: any, tails: double, type: double): double
        usDollar(arg1: double, arg2: double): string
        var_P(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        varP(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        vdb(cost: double, salvage: double, life: double, start_period: double, end_period: double, factor: any, no_switch: any): double
        vLookup(lookup_value: any, table_array: any, col_index_num: any, range_lookup: any): any
        weekday(Serial_number: any, return_type: any): double
        weekNum(Serial_number: any, return_type: any): double
        weibull(x: double, alpha: double, beta: double, cumulative: boolean): double
        workday(start_date: any, days: any, holidays: any): long
        xirr(values: any, dates: any, guess: any): double
        yearFrac(start_date: any, end_date: any, basis: any): double
        yieldMat(settlement: any, maturity: any, issue: any, rate: any, pr: any, basis: any): double
        z_Test(array: any, x: double, sigma: any): double
        zTest(array: any, x: double, sigma: any): double
        rank(number: double, range: excel.Range, order: any): double
        accrInt(issue: any, first_interest: any, settlement: any, rate: any, par: any, frequency: any, basis: any): double
        rate(nper: double, pmt: double, pv: double, fv: any, type: any, guess: any): double
        accrIntM(issue: any, maturity: any, rate: any, par: any, basis: any): double
        acosh(number: double): double
        amorLinc(cost: any, date_purchased: any, first_period: any, salvage: any, period: any, rate: any, basis: any): double
        asc(string: string): string
        asinh(number: double): double
        atanh(number: double): double
        aveDev(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        bahtText(arg1: double): string
        besselI(x: any, n: any): double
        besselJ(x: any, n: any): double
        besselK(x: any, n: any): double
        besselY(x: any, n: any): double
        beta_Inv(probability: double, alpha: double, beta: double, a: any, b: any): double
        betaDist(x: double, alpha: double, beta: double, a: any, b: any): double
        betaInv(probability: double, alpha: double, beta: double, a: any, b: any): double
        bin2Dec(number: any): string
        bin2Hex(number: any, places: any): string
        bin2Oct(number: any, places: any): string
        ceiling(number: double, significance: double): double
        chiDist(x: double, degrees_freedom: double): double
        chiInv(probability: double, degrees_freedom: double): double
        chiTest(actual_range: any, expected_range: any): double
        choose(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): any
        complex(real_num: any, i_number: any, arg3: any): string
        combin(number: double, number_chosen: double): double
        correl(range1: any, range2: any): double
        countA(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        countIf(arg1: excel.Range, arg2: any): double
        countIfs(arg1: excel.Range, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        coupDays(settlement: any, maturity: any, frequency: any, basis: any): double
        coupNcd(settlement: any, maturity: any, frequency: any, basis: any): double
        coupNum(settlement: any, maturity: any, frequency: any, basis: any): double
        coupPcd(settlement: any, maturity: any, frequency: any, basis: any): double
        covar(arg1: any, arg2: any): double
        cumIPmt(rate: any, nper: any, pv: any, start_period: any, end_period: any, type: any): double
        cumPrinc(rate: any, nper: any, pv: any, start_period: any, end_period: any, type: any): double
        dAverage(database: excel.Range, field: any, criteria: any): double
        days360(start_date: any, end_date: any, method: any): double
        dbcs(arg1: string): string
        dCount(database: excel.Range, field: any, criteria: any): double
        dCountA(database: excel.Range, field: any, criteria: any): double
        ddb(cost: double, salvage: double, life: double, period: double, factor: any): double
        dec2Bin(number: any, places: any): string
        dec2Hex(number: any, places: any): string
        dec2Oct(number: any, places: any): string
        degrees(angle: double): double
        devSq(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        dGet(database: excel.Range, field: any, criteria: any): any
        disc(settlement: any, maturity: any, pr: any, redemption: any, basis: any): double
        dMax(database: excel.Range, field: any, criteria: any): double
        dMin(database: excel.Range, field: any, criteria: any): double
        dollar(currencyNumber: double, decimalNumber: any): string
        dollarDe(fractional_dollar: any, fraction: any): double
        dollarFr(decimal_dollar: any, fraction: any): double
        dProduct(database: excel.Range, field: any, criteria: any): double
        dStDev(database: excel.Range, field: any, criteria: any): double
        dStDevp(database: excel.Range, field: any, criteria: any): double
        dSum(database: excel.Range, field: any, criteria: any): double
        duration(settlement: any, maturity: any, coupon: any, yld: any, frequency: any, basis: any): double
        dVar(database: excel.Range, field: any, criteria: any): double
        dVarp(database: excel.Range, field: any, criteria: any): double
        eDate(start_date: any, method: any): double
        effect(norminal_rate: any, npery: any): double
        eoMonth(start_date: any, method: any): double
        erf(lower_limit: any, upper_limit: any): double
        erfC(x: any): double
        even(number: double): double
        intercept(arg1: any, arg2: any): double
        getRangeString(array2: any, array1: any): string
        getRangeString(array: any): string
        processFormula(formulaStr: string): any
        amorDegrc(cost: any, date_purchased: any, first_period: any, salvage: any, period: any, rate: any, basis: any): double
        processFormulas(formulaStr: string, args: any[]): any
        beta_Dist(x: double, alpha: double, beta: double, cumulative: boolean, a: any, b: any): double
        binom_Dist(number_s: double, trials: double, probability_s: double, cumulative: boolean): double
        binomDist(number_s: double, trials: double, probability_s: double, cumulative: boolean): double
        confidence(alpha: double, standard_dev: double, size: double): double
        countBlank(range: excel.Range): double
        coupDayBs(settlement: any, maturity: any, frequency: any, basis: any): double
        coupDaysNc(settlement: any, maturity: any, frequency: any, basis: any): double
        critBinom(trials: double, probability_s: double, alpha: double): double
        erf_Precise(lower_limit: any): double
        erfC_Precise(x: any): double
        expon_Dist(x: double, lambda: double, cumulative: boolean): double
        exponDist(x: double, lambda: double, cumulative: boolean): double
        f_Dist_RT(x: double, degrees_freedom1: double, degrees_freedom2: double): double
        factDouble(number: any): double
        fisherInv(y: double): double
        fVSchedule(principal: any, schedule: any): double
        gamma_Dist(x: double, alpha: double, beta: double, cumulative: boolean): double
        gamma_Inv(probability: double, alpha: double, beta: double): double
        gammaDist(x: double, alpha: double, beta: double, cumulative: boolean): double
        gammaLn_Precise(x: double): double
        hypGeomDist(sample_s: double, number_sample: double, population_s: double, number_population: double): double
        hypGeom_Dist(sample_s: double, number_sample: double, population_s: double, number_population: double, b: boolean): double
        imaginary(inumber: any): double
        imArgument(inumber: any): string
        imConjugate(inumber: any): string
        imProduct(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): string
        isLogical(arg1: any): boolean
        isNonText(arg1: any): boolean
        logNorm_Inv(probability: double, mean: double, standard_dev: double): double
        logNormDist(probability: double, mean: double, standard_dev: double): double
        getPromptString(type: int): string
        transpose(range: any): any
        mDuration(settlement: any, maturity: any, coupon: any, yld: any, frequency: any, basis: any): double
        mode_Sngl(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        multiNomial(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        negBinomDist(number_f: double, number_s: double, probability_s: double): double
        networkdays(start_date: any, end_date: any, holidays: any): double
        norm_Dist(x: double, mean: double, standard_dev: double, cumulative: boolean): double
        normSDist(z: double): double
        oddFPrice(settlement: any, maturity: any, issue: any, first_coupon: any, rate: any, yld: any, redemption: any, frequency: any, basis: any): double
        oddFYield(settlement: any, maturity: any, issue: any, first_coupon: any, rate: any, pr: any, redemption: any, frequency: any, basis: any): double
        oddLPrice(settlement: any, maturity: any, last_interest: any, rate: any, yld: any, redemption: any, frequency: any, basis: any): double
        oddLYield(settlement: any, maturity: any, last_interest: any, rate: any, pr: any, redemption: any, frequency: any, basis: any): double
        percentile(range: any, k: double): double
        percentile_Exc(range: any, k: double): double
        percentile_Inc(range: any, k: double): double
        percentRank(array: any, x: double, significance: any): double
        percentRank_Exc(array: any, x: double, significance: any): double
        percentRank_Inc(array: any, x: double, significance: any): double
        poisson_Dist(x: double, mean: double, cumulative: boolean): double
        priceDisc(settlement: any, maturity: any, discount: any, redemption: any, basis: any): double
        quartile_Exc(range: any, quart: double): double
        quartile_Inc(range: any, quart: double): double
        randBetween(bottom: any, top: any): double
        roundDown(number: double, digits: double): double
        seriesSum(x: any, n: any, m: any, coefficients: any): double
        standardize(x: double, mean: double, standard_dev: double): double
        substitute(originString: string, oldString: string, newString: string, orderNumber: any): string
        sumProduct(arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any, arg11: any, arg12: any, arg13: any, arg14: any, arg15: any, arg16: any, arg17: any, arg18: any, arg19: any, arg20: any, arg21: any, arg22: any, arg23: any, arg24: any, arg25: any, arg26: any, arg27: any, arg28: any, arg29: any, arg30: any): double
        t_Dist_2T(x: double, degrees_freedom: double): double
        t_Dist_RT(x: double, degrees_freedom: double): double
        tBillPrice(settlement: any, maturity: any, discount: any): double
        tBillYield(settlement: any, maturity: any, pr: any): double
        weibull_Dist(x: double, alpha: double, beta: double, cumulative: boolean): double
        yieldDisc(settlement: any, maturity: any, pr: any, redemption: any, basis: any): double
        getArrayString(values: string[], needBrackets: boolean): string
        getArrayString(values: any): string
        getArrayString(values: number[], needBrackets: boolean): string
        getArrayString(values: number[]): string
        getArrayString(values: number[], needBrackets: boolean): string
        getArrayString(values: string[]): string
        getArrayString(values: any): string
        getArrayString(values: number[]): string
        getArrayString(values: any): string
        getArrayString(values: number[], needBrackets: boolean): string
        getArrayString(values: number[]): string
        getArrayString(values: any): string
        getArrayString(values: number[], needBrackets: boolean): string
        getArrayString(values: number[]): string
        getArrayString(values: any): string
        isFUNCTION_ERROR(str: string): boolean
    }
    export interface Worksheets {
        add(before: any, after: any, count: any, type: any): any
        delete(): void
        copy(before: any, after: any): void
        item(index: any): any
        getCount(): int
        getItem(index: any): any
        printPreview(enableChanges: any): void
        setVisible(pVisible: any): void
        move(before: any, after: any): void
        printOut(from: any, to: any, copies: any, preview: any, activePrinter: any, printToFile: any, collate: any, prToFileName: any, ignorePrintAreas: any): void
        select(replace: any): void
        getVisible(): any
        getWorksheet(msheet: any): excel.Worksheet
        getActiveSheet(): excel.Worksheet
        getHPageBreaks(): excel.HPageBreaks
        getVPageBreaks(): excel.VPageBreaks
        createWorksheet(msheet: any): excel.Worksheet
        fillAcrossSheets(mWorksheet: any[], range: excel.Range, type: int): void
        fillAcrossSheets(range: excel.Range, type: int): void
    }
    export interface Worksheets {
        add(before: any, after: any, count: any, type: any): any
        delete(): void
        copy(before: any, after: any): void
        item(index: any): any
        getCount(): int
        getItem(index: any): any
        printPreview(enableChanges: any): void
        setVisible(pVisible: any): void
        move(before: any, after: any): void
        printOut(from: any, to: any, copies: any, preview: any, activePrinter: any, printToFile: any, collate: any, prToFileName: any, ignorePrintAreas: any): void
        select(replace: any): void
        getVisible(): any
        getWorksheet(msheet: any): excel.Worksheet
        getActiveSheet(): excel.Worksheet
        getHPageBreaks(): excel.HPageBreaks
        getVPageBreaks(): excel.VPageBreaks
        createWorksheet(msheet: any): excel.Worksheet
        fillAcrossSheets(mWorksheet: any[], range: excel.Range, type: int): void
        fillAcrossSheets(range: excel.Range, type: int): void
    }
    export interface WorksheetView {
        getSheet(): excel.Worksheet
        setDisplayZeros(display: boolean): void
        isDisplayZeros(): boolean
        setDisplayFormulas(display: boolean): void
        isDisplayFormulas(): boolean
        setDisplayGridlines(display: boolean): void
        isDisplayGridlines(): boolean
        setDisplayHeadings(display: boolean): void
        isDisplayHeadings(): boolean
    }
    export interface WorksheetView {
        getSheet(): excel.Worksheet
        setDisplayZeros(display: boolean): void
        isDisplayZeros(): boolean
        setDisplayFormulas(display: boolean): void
        isDisplayFormulas(): boolean
        setDisplayGridlines(display: boolean): void
        isDisplayGridlines(): boolean
        setDisplayHeadings(display: boolean): void
        isDisplayHeadings(): boolean
    }
    export interface XmlDataBinding {
        refresh(): int
        getSourceUrl(): string
        clearSettings(): void
        LoadSettings(url: string): void
    }
    export interface XmlDataBinding {
        refresh(): int
        getSourceUrl(): string
        clearSettings(): void
        LoadSettings(url: string): void
    }
    export interface XmlMap {
        getName(): string
        delete(): void
        setName(pName: string): void
        export(url: string, overwrite: boolean): int
        import1(url: string, overwrite: any): int
        isShowImportExportValidationErrors(): boolean
        setShowImportExportValidationErrors(pShowImportExportValidationErrors: boolean): void
        isAdjustColumnWidth(): boolean
        setAdjustColumnWidth(pAdjustColumnWidth: boolean): void
        isPreserveColumnFilter(): boolean
        getRootElementName(): string
        getRootElementNamespace(): excel.XmlNamespace
        setAppendOnImport(pAppendOnImport: boolean): void
        setPreserveColumnFilter(pPreserveColumnFilter: boolean): void
        isPreserveNumberFormatting(): boolean
        isSaveDataSourceDefinition(): boolean
        setPreserveNumberFormatting(pPreserveNumberFormatting: boolean): void
        setSaveDataSourceDefinition(pSaveDataSourceDefinition: boolean): void
        isAppendOnImport(): boolean
        getDataBinding(): excel.XmlDataBinding
        isExportable(): boolean
        getSchemas(): excel.XmlSchemas
        exportXml(data: string): int
        importXml(xmlData: string, overwrite: any): int
    }
    export interface XmlMap {
        getName(): string
        delete(): void
        setName(pName: string): void
        export(url: string, overwrite: boolean): int
        import1(url: string, overwrite: any): int
        isShowImportExportValidationErrors(): boolean
        setShowImportExportValidationErrors(pShowImportExportValidationErrors: boolean): void
        isAdjustColumnWidth(): boolean
        setAdjustColumnWidth(pAdjustColumnWidth: boolean): void
        isPreserveColumnFilter(): boolean
        getRootElementName(): string
        getRootElementNamespace(): excel.XmlNamespace
        setAppendOnImport(pAppendOnImport: boolean): void
        setPreserveColumnFilter(pPreserveColumnFilter: boolean): void
        isPreserveNumberFormatting(): boolean
        isSaveDataSourceDefinition(): boolean
        setPreserveNumberFormatting(pPreserveNumberFormatting: boolean): void
        setSaveDataSourceDefinition(pSaveDataSourceDefinition: boolean): void
        isAppendOnImport(): boolean
        getDataBinding(): excel.XmlDataBinding
        isExportable(): boolean
        getSchemas(): excel.XmlSchemas
        exportXml(data: string): int
        importXml(xmlData: string, overwrite: any): int
    }
    export interface XmlMaps {
        add(schema: string, rootElementName: string): excel.XmlMap
        item(index: any): excel.XmlMap
        getCount(): int
    }
    export interface XmlMaps {
        add(schema: string, rootElementName: string): excel.XmlMap
        item(index: any): excel.XmlMap
        getCount(): int
    }
    export interface XmlNamespace {
        getUri(): string
        getPrefix(): string
    }
    export interface XmlNamespace {
        getUri(): string
        getPrefix(): string
    }
    export interface XmlNamespaces {
        getValue(): string
        item(index: any): excel.XmlNamespace
        getCount(): int
        InstallManifest(path: string, installForAllUsers: boolean): void
    }
    export interface XmlNamespaces {
        getValue(): string
        item(index: any): excel.XmlNamespace
        getCount(): int
        InstallManifest(path: string, installForAllUsers: boolean): void
    }
    export interface XmlSchema {
        getName(): string
        getXML(): string
        getNamespace(): excel.XmlNamespace
    }
    export interface XmlSchema {
        getName(): string
        getXML(): string
        getNamespace(): excel.XmlNamespace
    }
    export interface XmlSchemas {
        item(index: any): excel.XmlSchema
        getCount(): int
    }
    export interface XmlSchemas {
        item(index: any): excel.XmlSchema
        getCount(): int
    }
    export interface XPath {
        clear(): void
        getValue(): string
        setValue(map: excel.XmlMap, xPath: string, selectionNamespace: any, repeating: boolean): void
        getMap(): excel.XmlMap
        isRepeating(): boolean
    }
    export interface XPath {
        clear(): void
        getValue(): string
        setValue(map: excel.XmlMap, xPath: string, selectionNamespace: any, repeating: boolean): void
        getMap(): excel.XmlMap
        isRepeating(): boolean
    }
}

declare namespace excel.constants {
    export const enum ChartConstants {
        EMOCHART_COLUMN = 0,
        EMOCOLUMN_CLUSTERED = 0,
        EMOCOLUMN_STACKED = 1,
        EMOCOLUMN_STACKED100 = 2,
        EMO3DCOLUMN_CLUSTERED = 3,
        EMO3DCOLUMN_STACKED = 4,
        EMO3DCOLUMN_STACKED100 = 5,
        EMO3DCOLUMN_3D = 6,
        EMOCHART_BAR = 1,
        EMOBAR_CLUSTERED = 0,
        EMOBAR_STACKED = 1,
        EMOBAR_STACKED100 = 2,
        EMO3DBAR_CLUSTERED = 3,
        EMO3DBAR_STACKED = 4,
        EMO3DBAR_STACKED100 = 5,
        EMOCHART_LINE = 2,
        EMOLINE = 0,
        EMOLINE_STACKED = 1,
        EMOLINE_STACKED100 = 2,
        EMOLINE_MARKERS = 3,
        EMOLINE_MARKERS_STACKED = 4,
        EMOLINE_MARKERS_STACKED100 = 5,
        EMOLINE_3D = 6,
        EMOCHART_PIE = 3,
        EMOPIE = 0,
        EMO3DPIE = 1,
        EMOPIEOFPIE = 2,
        EMO3DPIE_EXPLODED = 3,
        EMOPIE_EXPLODED = 4,
        EMOBAROFPIE = 5,
        EMOCHART_XYSCATTER = 4,
        EMOXYSCATTER = 0,
        EMOXYSCATTER_LINES = 1,
        EMOXYSCATTER_LINESNOMARKERS = 2,
        EMOXYSCATTER_SMOOTH = 3,
        EMOXYSCATTER_SMOOTHNOMARKERS = 4,
        EMOCHART_AREA = 5,
        EMOAREA = 0,
        EMOAREA_STACKED = 1,
        EMOAREA_STACKED100 = 2,
        EMOAREA_3DAREA = 3,
        EMOAREA_3DSTACKED = 4,
        EMOAREA_3DSTACKED100 = 5,
        EMOCHART_DOUGHNUT = 6,
        EMODOUGHNUT = 0,
        EMOOUGHNUT_EXPLODED = 1,
        EMOCHART_RADAR = 7,
        EMORADAR = 0,
        EMORADAR_MARKERS = 1,
        EMORADAR_FILLED = 2,
        EMOCHART_BUBBLE = 8,
        EMOBUBBLE = 0,
        EMOBUBBLE_3D = 1,
        EMOCHART_STOCK = 9,
        EMOSTOCK_HLC = 0,
        EMOSTOCK_OHLC = 1,
        EMOSTOCK_VHLC = 2,
        EMOSTOCK_VOHLC = 3,
        EMOCHART_CYLINDER = 10,
        EMOCYLINDER_COLCLUSTERED = 0,
        EMOCYLINDER_STACKED = 1,
        EMOCYLINDER_COLSTACKED100 = 2,
        EMOCYLINDER_BARCLUSTERED = 3,
        EMOCYLINDER_BARSTACKED = 4,
        EMOCYLINDER_BARSTACKED100 = 5,
        EMOCYLINDER_COL = 6,
        EMOCHART_CONE = 11,
        EMOCONE_COLCLUSTERED = 0,
        EMOCONE_COLSTACKED = 1,
        EMOCONE_COLCLUSTERED100 = 2,
        EMOCONE_BARCLUSTERED = 3,
        EMOCONE_BARSTACKED = 4,
        EMOCONE_BARCLUSTERED100 = 5,
        EMOCONE_COL = 6,
        EMOCHART_PYRAMID = 12,
        EMOPYRAMID_COLCLUSTERED = 0,
        EMOPYRAMID_COLSTACKED = 1,
        EMOPYRAMID_COLCLUSTERED100 = 2,
        EMOPYRAMID_BARCLUSTERED = 3,
        EMOPYRAMID_BARSTACKED = 4,
        EMOPYRAMID_BARCLUSTERED100 = 5,
        EMOPYRAMID_COL = 6,
        EMOCHART_CURVE = 13,
        EMOCURVE_SURFACE = 0,
        EMOCURVE_SURFACEWIREFRAME = 1,
        EMOCURVE_SURFACETOPVIEW = 2,
        EMOCURVE_SURFACETOPVIEWWIREFRAME = 3,
        BARAREA_STACKED = 0,
        BARAREA_STACKED100 = 1,
        BARBAR_ClUSTERED = 2,
        BA3DRBAR_ClUSTERED = 3,
        BARCOLUMN_ClUSTERED = 4,
        BAR3DCOLUMN = 5,
        BAR_RADAR = 6,
        BAR_BUBBLE = 7,
        BAR_LINEMARKERS = 8,
        BAR_XYSCATTER = 9,
        BAR_PIE = 10,
        BAR_DOUGHNUT = 11,
        SERIES_TYPE_NONE = 0,
        SERIES_TYPE_ROW = 1,
        SERIES_TYPE_COL = 2,
        LEGEND_NONE = 0,
        LEGEND_BOTTOM = 1,
        LEGEND_TOPRIGHT = 2,
        LEGEND_TOP = 3,
        LEGEND_RIGHT = 4,
        LEGEND_LEFT = 5,
        CHART_TITLE = 0,
        X_TITLE = 1,
        Y_TITLE = 2,
        XSCALE_TITLE = 3,
        YSCALE_TITLE = 4,
        SECOND_Y_TITLE = 5,
        LINESTYLE_NONE = 0,
        LINESTYLE_SOLID = 1,
        LINESTYLE_LLS_DASH = 2,
        LINESTYLE_SHORT_DASH = 3,
        LINESTYLE_SSL_DASH = 4,
        LINESTYLE_LSL_DASH = 5,
        LINESTYLE_DOT_50 = 6,
        LINESTYLE_DOT_70 = 7,
        LINESTYLE_DOT_25 = 8,
        LINESTYLE_CUSTOM = 9,
        BORDER_AUTO = 0,
        BORDER_NONE = 1,
        BORDER_CUSTOM = 2,
        FIRST_WIDTH = 1,
        SECOND_WIDTH = 2,
        THIRD_WIDTH = 3,
        FOUR_WIDTH = 4,
        THREADLINE_LINETYPE = 0,
        THREADLINE_LOGARITHM = 1,
        THREADLINE_GUANTIC = 2,
        THREADLINE_POWER = 3,
        THREADLINE_EXPONENT = 4,
        THREADLINE_HORDE = 5,
        FIRST_X_AXIS = 0,
        SECOND_X_AXIS = 1,
        FIRST_Y_AXIS = 0,
        SECOND_Y_AXIS = 1,
        CATEGORY_AXIS = 0,
        SECOND_CATEGORY_AXIS = 2,
        VALUE_AXIS = 1,
        SECOND_VALUE_AXIS = 3,
        SCALE_UNIT_NONE = 0,
        SCALE_UNIT_BAI = 1,
        SCALE_UNIT_QIAN = 2,
        SCALE_UNIT_WAN = 3,
        SCALE_UNIT_SHIWAN = 4,
        SCALE_UNIT_BAIWAN = 5,
        SCALE_UNIT_QIANWAN = 6,
        SCALE_UNIT_YI = 7,
        SCALE_UNIT_SHIYI = 8,
        SCALE_UNIT_ZHAO = 9,
        FILL_AUTO = 0,
        FILL_NONE = 1,
        FILL_CUSTOM = 2,
        ALIGNMENT_VER_UP = 0,
        ALIGNMENT_VER_MIDDLE = 1,
        ALIGNMENT_VER_DOWN = 2,
        ALIGNMENT_BOTH = 3,
        ALIGNMENT_SCATTER = 4,
        ALIGNMENT_HOR_LEFT = 0,
        ALIGNMENT_HOR_MIDDLE = 1,
        ALIGNMENT_HOR_RIGHT = 2,
        SCALE_NONE = 0,
        SCALE_OUTER = 1,
        SCALE_INNER = 2,
        SCALE_CROSS = 3,
        SCALE_LABEL_NONE = 0,
        SCALE_LABEL_TOP = 1,
        SCALE_LABEL_BOTTOM = 2,
        SCALE_LABEL_AXIS_SIDE = 3,
        LOCATION_FOLLOW = 0,
        LOCATION_NOT_FOLLOW = 1,
        LOCATION_NONE_FOLLOW = 2,
        MGRID_Y_VALUE = 0,
        MGRID_X_CATEGORY = 2,
        MGRID_SUB_Y_VALUE = 1,
        MGRID_SUB_X_CATEGORY = 3,
        DATA_TABLE_NONE = 0,
        DATA_TABLE_ONLY = 1,
        DATA_TABLE_BOTH = 2,
        NUMBERFORMAT_GENERAL = 0,
        NUMBERFORMAT_NUMBER = 1,
        NUMBERFORMAT_CURRENCY = 2,
        NUMBERFORMAT_ACCOUNTING = 3,
        NUMBERFORMAT_DATE = 4,
        NUMBERFORMAT_TIME = 5,
        NUMBERFORMAT_PERCENTAGE = 6,
        NUMBERFORMAT_FRACTION = 7,
        NUMBERFORMAT_SCIENTIFIC = 8,
        NUMBERFORMAT_TEXT = 9,
        NUMBERFORMAT_SPECIAL = 10,
        NUMBERFORMAT_CUSTOM = 11,
        NONE = 0,
        CIRCLE = 1,
        BOX = 2,
        TRIANGLE = 3,
        DIAMOND = 4,
        STAR = 5,
        VERT_LINE = 6,
        HORIZ_LINE = 7,
        CROSS_1 = 8,
        CROSS_2 = 9,
        LAST = 8,
        HEADFOOT_TOP = 0,
        HEADFOOT_EAST = 1,
        HEADFOOT_WEST = 2,
        HEADFOOT_SOUTH = 3,
        HEADFOOT_NORTH = 4,
        EMOCHART_3DAREA = 1,
        EMOCHART_3DBAR = 3,
        EMOCHART_3DCOLUMN = 5,
        EMOCHART_SCATTER_PLOT = 8,
        EMOCHART_SCATTE_POINT = 9,
        LEGEND_NORTH = 3,
        LEGEND_SOUTH = 1,
        LEGEND_EAST = 4,
        LEGEND_WEST = 5,
        LEGEND_NORTHEAST = 2,
        LEGEND_NORTHWEST = 6,
        LEGEND_SOUTHEAST = 8,
        LEGEND_SOUTHWEST = 7,
        LEGEND_HORIZONTAL = 0,
        LEGEND_VERTICAL = 1,
        FILLSTYLE_AUTO = 0,
        FILLSTYLE_NONE = 1,
        FILLSTYLE_CUSTOMIZE = 2,
        LINESTYLE_AUTO = 0,
        LINESTYLE_NO = 1,
        LINESTYLE_USERDEFINE = 2,
        FIRST_AXIS = 0,
        SECOND_AXIS = 1,
        AXIS_AUTOMATIC = 0,
        AXIS_CATEGORY = 1,
        AXIS_TIMESCALE = 2,
        GRIDLINE_MAJOR = 0,
        GRIDLINE_MINOR = 1

    }
    export const enum ChartConstants {
        EMOCHART_COLUMN = 0,
        EMOCOLUMN_CLUSTERED = 0,
        EMOCOLUMN_STACKED = 1,
        EMOCOLUMN_STACKED100 = 2,
        EMO3DCOLUMN_CLUSTERED = 3,
        EMO3DCOLUMN_STACKED = 4,
        EMO3DCOLUMN_STACKED100 = 5,
        EMO3DCOLUMN_3D = 6,
        EMOCHART_BAR = 1,
        EMOBAR_CLUSTERED = 0,
        EMOBAR_STACKED = 1,
        EMOBAR_STACKED100 = 2,
        EMO3DBAR_CLUSTERED = 3,
        EMO3DBAR_STACKED = 4,
        EMO3DBAR_STACKED100 = 5,
        EMOCHART_LINE = 2,
        EMOLINE = 0,
        EMOLINE_STACKED = 1,
        EMOLINE_STACKED100 = 2,
        EMOLINE_MARKERS = 3,
        EMOLINE_MARKERS_STACKED = 4,
        EMOLINE_MARKERS_STACKED100 = 5,
        EMOLINE_3D = 6,
        EMOCHART_PIE = 3,
        EMOPIE = 0,
        EMO3DPIE = 1,
        EMOPIEOFPIE = 2,
        EMO3DPIE_EXPLODED = 3,
        EMOPIE_EXPLODED = 4,
        EMOBAROFPIE = 5,
        EMOCHART_XYSCATTER = 4,
        EMOXYSCATTER = 0,
        EMOXYSCATTER_LINES = 1,
        EMOXYSCATTER_LINESNOMARKERS = 2,
        EMOXYSCATTER_SMOOTH = 3,
        EMOXYSCATTER_SMOOTHNOMARKERS = 4,
        EMOCHART_AREA = 5,
        EMOAREA = 0,
        EMOAREA_STACKED = 1,
        EMOAREA_STACKED100 = 2,
        EMOAREA_3DAREA = 3,
        EMOAREA_3DSTACKED = 4,
        EMOAREA_3DSTACKED100 = 5,
        EMOCHART_DOUGHNUT = 6,
        EMODOUGHNUT = 0,
        EMOOUGHNUT_EXPLODED = 1,
        EMOCHART_RADAR = 7,
        EMORADAR = 0,
        EMORADAR_MARKERS = 1,
        EMORADAR_FILLED = 2,
        EMOCHART_BUBBLE = 8,
        EMOBUBBLE = 0,
        EMOBUBBLE_3D = 1,
        EMOCHART_STOCK = 9,
        EMOSTOCK_HLC = 0,
        EMOSTOCK_OHLC = 1,
        EMOSTOCK_VHLC = 2,
        EMOSTOCK_VOHLC = 3,
        EMOCHART_CYLINDER = 10,
        EMOCYLINDER_COLCLUSTERED = 0,
        EMOCYLINDER_STACKED = 1,
        EMOCYLINDER_COLSTACKED100 = 2,
        EMOCYLINDER_BARCLUSTERED = 3,
        EMOCYLINDER_BARSTACKED = 4,
        EMOCYLINDER_BARSTACKED100 = 5,
        EMOCYLINDER_COL = 6,
        EMOCHART_CONE = 11,
        EMOCONE_COLCLUSTERED = 0,
        EMOCONE_COLSTACKED = 1,
        EMOCONE_COLCLUSTERED100 = 2,
        EMOCONE_BARCLUSTERED = 3,
        EMOCONE_BARSTACKED = 4,
        EMOCONE_BARCLUSTERED100 = 5,
        EMOCONE_COL = 6,
        EMOCHART_PYRAMID = 12,
        EMOPYRAMID_COLCLUSTERED = 0,
        EMOPYRAMID_COLSTACKED = 1,
        EMOPYRAMID_COLCLUSTERED100 = 2,
        EMOPYRAMID_BARCLUSTERED = 3,
        EMOPYRAMID_BARSTACKED = 4,
        EMOPYRAMID_BARCLUSTERED100 = 5,
        EMOPYRAMID_COL = 6,
        EMOCHART_CURVE = 13,
        EMOCURVE_SURFACE = 0,
        EMOCURVE_SURFACEWIREFRAME = 1,
        EMOCURVE_SURFACETOPVIEW = 2,
        EMOCURVE_SURFACETOPVIEWWIREFRAME = 3,
        BARAREA_STACKED = 0,
        BARAREA_STACKED100 = 1,
        BARBAR_ClUSTERED = 2,
        BA3DRBAR_ClUSTERED = 3,
        BARCOLUMN_ClUSTERED = 4,
        BAR3DCOLUMN = 5,
        BAR_RADAR = 6,
        BAR_BUBBLE = 7,
        BAR_LINEMARKERS = 8,
        BAR_XYSCATTER = 9,
        BAR_PIE = 10,
        BAR_DOUGHNUT = 11,
        SERIES_TYPE_NONE = 0,
        SERIES_TYPE_ROW = 1,
        SERIES_TYPE_COL = 2,
        LEGEND_NONE = 0,
        LEGEND_BOTTOM = 1,
        LEGEND_TOPRIGHT = 2,
        LEGEND_TOP = 3,
        LEGEND_RIGHT = 4,
        LEGEND_LEFT = 5,
        CHART_TITLE = 0,
        X_TITLE = 1,
        Y_TITLE = 2,
        XSCALE_TITLE = 3,
        YSCALE_TITLE = 4,
        SECOND_Y_TITLE = 5,
        LINESTYLE_NONE = 0,
        LINESTYLE_SOLID = 1,
        LINESTYLE_LLS_DASH = 2,
        LINESTYLE_SHORT_DASH = 3,
        LINESTYLE_SSL_DASH = 4,
        LINESTYLE_LSL_DASH = 5,
        LINESTYLE_DOT_50 = 6,
        LINESTYLE_DOT_70 = 7,
        LINESTYLE_DOT_25 = 8,
        LINESTYLE_CUSTOM = 9,
        BORDER_AUTO = 0,
        BORDER_NONE = 1,
        BORDER_CUSTOM = 2,
        FIRST_WIDTH = 1,
        SECOND_WIDTH = 2,
        THIRD_WIDTH = 3,
        FOUR_WIDTH = 4,
        THREADLINE_LINETYPE = 0,
        THREADLINE_LOGARITHM = 1,
        THREADLINE_GUANTIC = 2,
        THREADLINE_POWER = 3,
        THREADLINE_EXPONENT = 4,
        THREADLINE_HORDE = 5,
        FIRST_X_AXIS = 0,
        SECOND_X_AXIS = 1,
        FIRST_Y_AXIS = 0,
        SECOND_Y_AXIS = 1,
        CATEGORY_AXIS = 0,
        SECOND_CATEGORY_AXIS = 2,
        VALUE_AXIS = 1,
        SECOND_VALUE_AXIS = 3,
        SCALE_UNIT_NONE = 0,
        SCALE_UNIT_BAI = 1,
        SCALE_UNIT_QIAN = 2,
        SCALE_UNIT_WAN = 3,
        SCALE_UNIT_SHIWAN = 4,
        SCALE_UNIT_BAIWAN = 5,
        SCALE_UNIT_QIANWAN = 6,
        SCALE_UNIT_YI = 7,
        SCALE_UNIT_SHIYI = 8,
        SCALE_UNIT_ZHAO = 9,
        FILL_AUTO = 0,
        FILL_NONE = 1,
        FILL_CUSTOM = 2,
        ALIGNMENT_VER_UP = 0,
        ALIGNMENT_VER_MIDDLE = 1,
        ALIGNMENT_VER_DOWN = 2,
        ALIGNMENT_BOTH = 3,
        ALIGNMENT_SCATTER = 4,
        ALIGNMENT_HOR_LEFT = 0,
        ALIGNMENT_HOR_MIDDLE = 1,
        ALIGNMENT_HOR_RIGHT = 2,
        SCALE_NONE = 0,
        SCALE_OUTER = 1,
        SCALE_INNER = 2,
        SCALE_CROSS = 3,
        SCALE_LABEL_NONE = 0,
        SCALE_LABEL_TOP = 1,
        SCALE_LABEL_BOTTOM = 2,
        SCALE_LABEL_AXIS_SIDE = 3,
        LOCATION_FOLLOW = 0,
        LOCATION_NOT_FOLLOW = 1,
        LOCATION_NONE_FOLLOW = 2,
        MGRID_Y_VALUE = 0,
        MGRID_X_CATEGORY = 2,
        MGRID_SUB_Y_VALUE = 1,
        MGRID_SUB_X_CATEGORY = 3,
        DATA_TABLE_NONE = 0,
        DATA_TABLE_ONLY = 1,
        DATA_TABLE_BOTH = 2,
        NUMBERFORMAT_GENERAL = 0,
        NUMBERFORMAT_NUMBER = 1,
        NUMBERFORMAT_CURRENCY = 2,
        NUMBERFORMAT_ACCOUNTING = 3,
        NUMBERFORMAT_DATE = 4,
        NUMBERFORMAT_TIME = 5,
        NUMBERFORMAT_PERCENTAGE = 6,
        NUMBERFORMAT_FRACTION = 7,
        NUMBERFORMAT_SCIENTIFIC = 8,
        NUMBERFORMAT_TEXT = 9,
        NUMBERFORMAT_SPECIAL = 10,
        NUMBERFORMAT_CUSTOM = 11,
        NONE = 0,
        CIRCLE = 1,
        BOX = 2,
        TRIANGLE = 3,
        DIAMOND = 4,
        STAR = 5,
        VERT_LINE = 6,
        HORIZ_LINE = 7,
        CROSS_1 = 8,
        CROSS_2 = 9,
        LAST = 8,
        HEADFOOT_TOP = 0,
        HEADFOOT_EAST = 1,
        HEADFOOT_WEST = 2,
        HEADFOOT_SOUTH = 3,
        HEADFOOT_NORTH = 4,
        EMOCHART_3DAREA = 1,
        EMOCHART_3DBAR = 3,
        EMOCHART_3DCOLUMN = 5,
        EMOCHART_SCATTER_PLOT = 8,
        EMOCHART_SCATTE_POINT = 9,
        LEGEND_NORTH = 3,
        LEGEND_SOUTH = 1,
        LEGEND_EAST = 4,
        LEGEND_WEST = 5,
        LEGEND_NORTHEAST = 2,
        LEGEND_NORTHWEST = 6,
        LEGEND_SOUTHEAST = 8,
        LEGEND_SOUTHWEST = 7,
        LEGEND_HORIZONTAL = 0,
        LEGEND_VERTICAL = 1,
        FILLSTYLE_AUTO = 0,
        FILLSTYLE_NONE = 1,
        FILLSTYLE_CUSTOMIZE = 2,
        LINESTYLE_AUTO = 0,
        LINESTYLE_NO = 1,
        LINESTYLE_USERDEFINE = 2,
        FIRST_AXIS = 0,
        SECOND_AXIS = 1,
        AXIS_AUTOMATIC = 0,
        AXIS_CATEGORY = 1,
        AXIS_TIMESCALE = 2,
        GRIDLINE_MAJOR = 0,
        GRIDLINE_MINOR = 1

    }
    export const enum YxAboveBelow {
        XlAboveAverage = 0,
        XlAboveStdDev = 4,
        XlBelowAverage = 1,
        XlBelowStdDev = 5,
        XlEqualAboveAverage = 2,
        XlEqualBelowAverage = 3

    }
    export const enum YxAboveBelow {
        XlAboveAverage = 0,
        XlAboveStdDev = 4,
        XlBelowAverage = 1,
        XlBelowStdDev = 5,
        XlEqualAboveAverage = 2,
        XlEqualBelowAverage = 3

    }
    export const enum YxApplyNamesOrder {
        yxColumnThenRow = 2,
        yxRowThenColumn = 1

    }
    export const enum YxApplyNamesOrder {
        yxColumnThenRow = 2,
        yxRowThenColumn = 1

    }
    export const enum YxArabicModes {
        yxArabicBothStrict = 3,
        yxArabicNone = 0,
        yxArabicStrictAlefHamza = 1,
        yxArabicStrictFinalYaa = 2

    }
    export const enum YxArabicModes {
        yxArabicBothStrict = 3,
        yxArabicNone = 0,
        yxArabicStrictAlefHamza = 1,
        yxArabicStrictFinalYaa = 2

    }
    export const enum YxArrangeStyle {
        yxArrangeStyleCascade = 7,
        yxArrangeStyleHorizontal = -4128,
        yxArrangeStyleTiled = 1,
        yxArrangeStyleVertical = -4166

    }
    export const enum YxArrangeStyle {
        yxArrangeStyleCascade = 7,
        yxArrangeStyleHorizontal = -4128,
        yxArrangeStyleTiled = 1,
        yxArrangeStyleVertical = -4166

    }
    export const enum YxAutoFillType {
        yxFillCopy = 1,
        yxFillDays = 5,
        yxFillDefault = 0,
        yxFillFormats = 3,
        yxFillMonths = 7,
        yxFillSeries = 2,
        yxFillValues = 4,
        yxFillWeekdays = 6,
        yxFillYears = 8,
        yxGrowthTrend = 10,
        yxLinearTrend = 9

    }
    export const enum YxAutoFillType {
        yxFillCopy = 1,
        yxFillDays = 5,
        yxFillDefault = 0,
        yxFillFormats = 3,
        yxFillMonths = 7,
        yxFillSeries = 2,
        yxFillValues = 4,
        yxFillWeekdays = 6,
        yxFillYears = 8,
        yxGrowthTrend = 10,
        yxLinearTrend = 9

    }
    export const enum YxAutoFilterOperator {
        yxAnd = 1,
        yxBottom10Items = 4,
        yxBottom10Percent = 6,
        yxFilterCellColor = 8,
        yxFilterDynamic = 11,
        yxFilterFontColor = 9,
        yxFilterIcon = 10,
        yxFilterValues = 7,
        yxOr = 2,
        yxTop10Items = 3,
        yxTop10Percent = 5

    }
    export const enum YxAutoFilterOperator {
        yxAnd = 1,
        yxBottom10Items = 4,
        yxBottom10Percent = 6,
        yxFilterCellColor = 8,
        yxFilterDynamic = 11,
        yxFilterFontColor = 9,
        yxFilterIcon = 10,
        yxFilterValues = 7,
        yxOr = 2,
        yxTop10Items = 3,
        yxTop10Percent = 5

    }
    export const enum YxAxisCrosses {
        yxAxisCrossesAutomatic = -4105,
        yxAxisCrossesCustom = -4114,
        yxAxisCrossesMaximum = 2,
        yxAxisCrossesMinimum = 4

    }
    export const enum YxAxisCrosses {
        yxAxisCrossesAutomatic = -4105,
        yxAxisCrossesCustom = -4114,
        yxAxisCrossesMaximum = 2,
        yxAxisCrossesMinimum = 4

    }
    export const enum YxAxisGroup {
        yxPrimary = 1,
        yxSecondary = 2

    }
    export const enum YxAxisGroup {
        yxPrimary = 1,
        yxSecondary = 2

    }
    export const enum YxAxisType {
        yxCategory = 1,
        yxSeriesAxis = 3,
        yxValue = 2

    }
    export const enum YxAxisType {
        yxCategory = 1,
        yxSeriesAxis = 3,
        yxValue = 2

    }
    export const enum YxBarShape {
        yxBox = 0,
        yxConeToMax = 5,
        yxConeToPoint = 4,
        yxCylinder = 3,
        yxPyramidToMax = 2,
        yxPyramidToPoint = 1

    }
    export const enum YxBarShape {
        yxBox = 0,
        yxConeToMax = 5,
        yxConeToPoint = 4,
        yxCylinder = 3,
        yxPyramidToMax = 2,
        yxPyramidToPoint = 1

    }
    export const enum YxBordersIndex {
        yxDiagonalDown = 5,
        yxDiagonalUp = 6,
        yxEdgeBottom = 9,
        yxEdgeLeft = 7,
        yxEdgeRight = 10,
        yxEdgeTop = 8,
        yxInsideHorizontal = 12,
        yxInsideVertical = 11

    }
    export const enum YxBordersIndex {
        yxDiagonalDown = 5,
        yxDiagonalUp = 6,
        yxEdgeBottom = 9,
        yxEdgeLeft = 7,
        yxEdgeRight = 10,
        yxEdgeTop = 8,
        yxInsideHorizontal = 12,
        yxInsideVertical = 11

    }
    export const enum YxBorderWeight {
        yxHairline = 1,
        yxMedium = -4138,
        yxThick = 4,
        yxThin = 2

    }
    export const enum YxBorderWeight {
        yxHairline = 1,
        yxMedium = -4138,
        yxThick = 4,
        yxThin = 2

    }
    export const enum YxBuiltInDialog {
        yxDialogActivate = 103,
        yxDialogActiveCellFont = 476,
        yxDialogAddChartAutoformat = 390,
        yxDialogAddinManager = 321,
        yxDialogAlignment = 43,
        yxDialogApplyNames = 133,
        yxDialogApplyStyle = 212,
        yxDialogAppMove = 170,
        yxDialogAppSize = 171,
        yxDialogArrangeAll = 12,
        yxDialogAssignToObject = 213,
        yxDialogAssignToTool = 293,
        yxDialogAttachText = 80,
        yxDialogAttachToolbars = 323,
        yxDialogAutoCorrect = 485,
        yxDialogAxes = 78,
        yxDialogBorder = 45,
        yxDialogCalculation = 32,
        yxDialogCellProtection = 46,
        yxDialogChangeLink = 166,
        yxDialogChartAddData = 392,
        yxDialogChartLocation = 527,
        yxDialogChartOptionsDataLabelMultiple = 724,
        yxDialogChartOptionsDataLabels = 505,
        yxDialogChartOptionsDataTable = 506,
        yxDialogChartSourceData = 540,
        yxDialogChartTrend = 350,
        yxDialogChartType = 526,
        yxDialogChartWizard = 288,
        yxDialogCheckboxProperties = 435,
        yxDialogClear = 52,
        yxDialogColorPalette = 161,
        yxDialogColumnWidth = 47,
        yxDialogCombination = 73,
        yxDialogConditionalFormatting = 583,
        yxDialogConsolidate = 191,
        yxDialogCopyChart = 147,
        yxDialogCopyPicture = 108,
        yxDialogCreateList = 796,
        yxDialogCreateNames = 62,
        yxDialogCreatePublisher = 217,
        yxDialogCustomizeToolbar = 276,
        yxDialogCustomViews = 493,
        yxDialogDataDelete = 36,
        yxDialogDataLabel = 379,
        yxDialogDataLabelMultiple = 723,
        yxDialogDataSeries = 40,
        yxDialogDataValidation = 525,
        yxDialogDefineName = 61,
        yxDialogDefineStyle = 229,
        yxDialogDeleteFormat = 111,
        yxDialogDeleteName = 110,
        yxDialogDemote = 203,
        yxDialogDisplay = 27,
        yxDialogDocumentInspector = 862,
        yxDialogEditboxProperties = 438,
        yxDialogEditColor = 223,
        yxDialogEditDelete = 54,
        yxDialogEditionOptions = 251,
        yxDialogEditSeries = 228,
        yxDialogErrorbarX = 463,
        yxDialogErrorbarY = 464,
        yxDialogErrorChecking = 732,
        yxDialogEvaluateFormula = 709,
        yxDialogExternalDataProperties = 530,
        yxDialogExtract = 35,
        yxDialogFileDelete = 6,
        yxDialogFileSharing = 481,
        yxDialogFillGroup = 200,
        yxDialogFillWorkgroup = 301,
        yxDialogFilter = 447,
        yxDialogFilterAdvanced = 370,
        yxDialogFindFile = 475,
        yxDialogFont = 26,
        yxDialogFontProperties = 381,
        yxDialogFormatAuto = 269,
        yxDialogFormatChart = 465,
        yxDialogFormatCharttype = 423,
        yxDialogFormatFont = 150,
        yxDialogFormatLegend = 88,
        yxDialogFormatMain = 225,
        yxDialogFormatMove = 128,
        yxDialogFormatNumber = 42,
        yxDialogFormatOverlay = 226,
        yxDialogFormatSize = 129,
        yxDialogFormatText = 89,
        yxDialogFormulaFind = 64,
        yxDialogFormulaGoto = 63,
        yxDialogFormulaReplace = 130,
        yxDialogFunctionWizard = 450,
        yxDialogGallery3dArea = 193,
        yxDialogGallery3dBar = 272,
        yxDialogGallery3dColumn = 194,
        yxDialogGallery3dLine = 195,
        yxDialogGallery3dPie = 196,
        yxDialogGallery3dSurface = 273,
        yxDialogGalleryArea = 67,
        yxDialogGalleryBar = 68,
        yxDialogGalleryColumn = 69,
        yxDialogGalleryCustom = 388,
        yxDialogGalleryDoughnut = 344,
        yxDialogGalleryLine = 70,
        yxDialogGalleryPie = 71,
        yxDialogGalleryRadar = 249,
        yxDialogGalleryScatter = 72,
        yxDialogGoalSeek = 198,
        yxDialogGridlines = 76,
        yxDialogImportTextFile = 666,
        yxDialogInsert = 55,
        yxDialogInsertHyperlink = 596,
        yxDialogInsertObject = 259,
        yxDialogInsertPicture = 342,
        yxDialogInsertTitle = 380,
        yxDialogLabelProperties = 436,
        yxDialogListboxProperties = 437,
        yxDialogMacroOptions = 382,
        yxDialogMailEditMailer = 470,
        yxDialogMailLogon = 339,
        yxDialogMailNextLetter = 378,
        yxDialogMainChart = 85,
        yxDialogMainChartType = 185,
        yxDialogMenuEditor = 322,
        yxDialogMove = 262,
        yxDialogMyPermission = 834,
        yxDialogNameManager = 977,
        yxDialogNew = 119,
        yxDialogNewName = 978,
        yxDialogNewWebQuery = 667,
        yxDialogNote = 154,
        yxDialogObjectProperties = 207,
        yxDialogObjectProtection = 214,
        yxDialogOpen = 1,
        yxDialogOpenLinks = 2,
        yxDialogOpenMail = 188,
        yxDialogOpenText = 441,
        yxDialogOptionsCalculation = 318,
        yxDialogOptionsChart = 325,
        yxDialogOptionsEdit = 319,
        yxDialogOptionsGeneral = 356,
        yxDialogOptionsListsAdd = 458,
        yxDialogOptionsME = 647,
        yxDialogOptionsTransition = 355,
        yxDialogOptionsView = 320,
        yxDialogOutline = 142,
        yxDialogOverlay = 86,
        yxDialogOverlayChartType = 186,
        yxDialogPageSetup = 7,
        yxDialogParse = 91,
        yxDialogPasteNames = 58,
        yxDialogPasteSpecial = 53,
        yxDialogPatterns = 84,
        yxDialogPermission = 832,
        yxDialogPhonetic = 656,
        yxDialogPivotCalculatedField = 570,
        yxDialogPivotCalculatedItem = 572,
        yxDialogPivotClientServerSet = 689,
        yxDialogPivotFieldGroup = 433,
        yxDialogPivotFieldProperties = 313,
        yxDialogPivotFieldUngroup = 434,
        yxDialogPivotShowPages = 421,
        yxDialogPivotSolveOrder = 568,
        yxDialogPivotTableOptions = 567,
        yxDialogPivotTableWizard = 312,
        yxDialogPlacement = 300,
        yxDialogPrint = 8,
        yxDialogPrinterSetup = 9,
        yxDialogPrintPreview = 222,
        yxDialogPromote = 202,
        yxDialogProperties = 474,
        yxDialogPropertyFields = 754,
        yxDialogProtectDocument = 28,
        yxDialogProtectSharing = 620,
        yxDialogPublishAsWebPage = 653,
        yxDialogPushbuttonProperties = 445,
        yxDialogReplaceFont = 134,
        yxDialogRoutingSlip = 336,
        yxDialogRowHeight = 127,
        yxDialogRun = 17,
        yxDialogSaveAs = 5,
        yxDialogSaveCopyAs = 456,
        yxDialogSaveNewObject = 208,
        yxDialogSaveWorkbook = 145,
        yxDialogSaveWorkspace = 285,
        yxDialogScale = 87,
        yxDialogScenarioAdd = 307,
        yxDialogScenarioCells = 305,
        yxDialogScenarioEdit = 308,
        yxDialogScenarioMerge = 473,
        yxDialogScenarioSummary = 311,
        yxDialogScrollbarProperties = 420,
        yxDialogSearch = 731,
        yxDialogSelectSpecial = 132,
        yxDialogSendMail = 189,
        yxDialogSeriesAxes = 460,
        yxDialogSeriesOptions = 557,
        yxDialogSeriesOrder = 466,
        yxDialogSeriesShape = 504,
        yxDialogSeriesX = 461,
        yxDialogSeriesY = 462,
        yxDialogSetBackgroundPicture = 509,
        yxDialogSetPrintTitles = 23,
        yxDialogSetUpdateStatus = 159,
        yxDialogShowDetail = 204,
        yxDialogShowToolbar = 220,
        yxDialogSize = 261,
        yxDialogSort = 39,
        yxDialogSortSpecial = 192,
        yxDialogSplit = 137,
        yxDialogStandardFont = 190,
        yxDialogStandardWidth = 472,
        yxDialogStyle = 44,
        yxDialogSubscribeTo = 218,
        yxDialogSubtotalCreate = 398,
        yxDialogSummaryInfo = 474,
        yxDialogTable = 41,
        yxDialogTabOrder = 394,
        yxDialogTextToColumns = 422,
        yxDialogUnhide = 94,
        yxDialogUpdateLink = 201,
        yxDialogVbaInsertFile = 328,
        yxDialogVbaMakeAddin = 478,
        yxDialogVbaProcedureDefinition = 330,
        yxDialogView3d = 197,
        yxDialogWebOptionsBrowsers = 773,
        yxDialogWebOptionsEncoding = 686,
        yxDialogWebOptionsFiles = 684,
        yxDialogWebOptionsFonts = 687,
        yxDialogWebOptionsGeneral = 683,
        yxDialogWebOptionsPictures = 685,
        yxDialogWindowMove = 14,
        yxDialogWindowSize = 13,
        yxDialogWorkbookAdd = 281,
        yxDialogWorkbookCopy = 283,
        yxDialogWorkbookInsert = 354,
        yxDialogWorkbookMove = 282,
        yxDialogWorkbookName = 386,
        yxDialogWorkbookNew = 302,
        yxDialogWorkbookOptions = 284,
        yxDialogWorkbookProtect = 417,
        yxDialogWorkbookTabSplit = 415,
        yxDialogWorkbookUnhide = 384,
        yxDialogWorkgroup = 199,
        yxDialogWorkspace = 95,
        yxDialogZoom = 256

    }
    export const enum YxBuiltInDialog {
        yxDialogActivate = 103,
        yxDialogActiveCellFont = 476,
        yxDialogAddChartAutoformat = 390,
        yxDialogAddinManager = 321,
        yxDialogAlignment = 43,
        yxDialogApplyNames = 133,
        yxDialogApplyStyle = 212,
        yxDialogAppMove = 170,
        yxDialogAppSize = 171,
        yxDialogArrangeAll = 12,
        yxDialogAssignToObject = 213,
        yxDialogAssignToTool = 293,
        yxDialogAttachText = 80,
        yxDialogAttachToolbars = 323,
        yxDialogAutoCorrect = 485,
        yxDialogAxes = 78,
        yxDialogBorder = 45,
        yxDialogCalculation = 32,
        yxDialogCellProtection = 46,
        yxDialogChangeLink = 166,
        yxDialogChartAddData = 392,
        yxDialogChartLocation = 527,
        yxDialogChartOptionsDataLabelMultiple = 724,
        yxDialogChartOptionsDataLabels = 505,
        yxDialogChartOptionsDataTable = 506,
        yxDialogChartSourceData = 540,
        yxDialogChartTrend = 350,
        yxDialogChartType = 526,
        yxDialogChartWizard = 288,
        yxDialogCheckboxProperties = 435,
        yxDialogClear = 52,
        yxDialogColorPalette = 161,
        yxDialogColumnWidth = 47,
        yxDialogCombination = 73,
        yxDialogConditionalFormatting = 583,
        yxDialogConsolidate = 191,
        yxDialogCopyChart = 147,
        yxDialogCopyPicture = 108,
        yxDialogCreateList = 796,
        yxDialogCreateNames = 62,
        yxDialogCreatePublisher = 217,
        yxDialogCustomizeToolbar = 276,
        yxDialogCustomViews = 493,
        yxDialogDataDelete = 36,
        yxDialogDataLabel = 379,
        yxDialogDataLabelMultiple = 723,
        yxDialogDataSeries = 40,
        yxDialogDataValidation = 525,
        yxDialogDefineName = 61,
        yxDialogDefineStyle = 229,
        yxDialogDeleteFormat = 111,
        yxDialogDeleteName = 110,
        yxDialogDemote = 203,
        yxDialogDisplay = 27,
        yxDialogDocumentInspector = 862,
        yxDialogEditboxProperties = 438,
        yxDialogEditColor = 223,
        yxDialogEditDelete = 54,
        yxDialogEditionOptions = 251,
        yxDialogEditSeries = 228,
        yxDialogErrorbarX = 463,
        yxDialogErrorbarY = 464,
        yxDialogErrorChecking = 732,
        yxDialogEvaluateFormula = 709,
        yxDialogExternalDataProperties = 530,
        yxDialogExtract = 35,
        yxDialogFileDelete = 6,
        yxDialogFileSharing = 481,
        yxDialogFillGroup = 200,
        yxDialogFillWorkgroup = 301,
        yxDialogFilter = 447,
        yxDialogFilterAdvanced = 370,
        yxDialogFindFile = 475,
        yxDialogFont = 26,
        yxDialogFontProperties = 381,
        yxDialogFormatAuto = 269,
        yxDialogFormatChart = 465,
        yxDialogFormatCharttype = 423,
        yxDialogFormatFont = 150,
        yxDialogFormatLegend = 88,
        yxDialogFormatMain = 225,
        yxDialogFormatMove = 128,
        yxDialogFormatNumber = 42,
        yxDialogFormatOverlay = 226,
        yxDialogFormatSize = 129,
        yxDialogFormatText = 89,
        yxDialogFormulaFind = 64,
        yxDialogFormulaGoto = 63,
        yxDialogFormulaReplace = 130,
        yxDialogFunctionWizard = 450,
        yxDialogGallery3dArea = 193,
        yxDialogGallery3dBar = 272,
        yxDialogGallery3dColumn = 194,
        yxDialogGallery3dLine = 195,
        yxDialogGallery3dPie = 196,
        yxDialogGallery3dSurface = 273,
        yxDialogGalleryArea = 67,
        yxDialogGalleryBar = 68,
        yxDialogGalleryColumn = 69,
        yxDialogGalleryCustom = 388,
        yxDialogGalleryDoughnut = 344,
        yxDialogGalleryLine = 70,
        yxDialogGalleryPie = 71,
        yxDialogGalleryRadar = 249,
        yxDialogGalleryScatter = 72,
        yxDialogGoalSeek = 198,
        yxDialogGridlines = 76,
        yxDialogImportTextFile = 666,
        yxDialogInsert = 55,
        yxDialogInsertHyperlink = 596,
        yxDialogInsertObject = 259,
        yxDialogInsertPicture = 342,
        yxDialogInsertTitle = 380,
        yxDialogLabelProperties = 436,
        yxDialogListboxProperties = 437,
        yxDialogMacroOptions = 382,
        yxDialogMailEditMailer = 470,
        yxDialogMailLogon = 339,
        yxDialogMailNextLetter = 378,
        yxDialogMainChart = 85,
        yxDialogMainChartType = 185,
        yxDialogMenuEditor = 322,
        yxDialogMove = 262,
        yxDialogMyPermission = 834,
        yxDialogNameManager = 977,
        yxDialogNew = 119,
        yxDialogNewName = 978,
        yxDialogNewWebQuery = 667,
        yxDialogNote = 154,
        yxDialogObjectProperties = 207,
        yxDialogObjectProtection = 214,
        yxDialogOpen = 1,
        yxDialogOpenLinks = 2,
        yxDialogOpenMail = 188,
        yxDialogOpenText = 441,
        yxDialogOptionsCalculation = 318,
        yxDialogOptionsChart = 325,
        yxDialogOptionsEdit = 319,
        yxDialogOptionsGeneral = 356,
        yxDialogOptionsListsAdd = 458,
        yxDialogOptionsME = 647,
        yxDialogOptionsTransition = 355,
        yxDialogOptionsView = 320,
        yxDialogOutline = 142,
        yxDialogOverlay = 86,
        yxDialogOverlayChartType = 186,
        yxDialogPageSetup = 7,
        yxDialogParse = 91,
        yxDialogPasteNames = 58,
        yxDialogPasteSpecial = 53,
        yxDialogPatterns = 84,
        yxDialogPermission = 832,
        yxDialogPhonetic = 656,
        yxDialogPivotCalculatedField = 570,
        yxDialogPivotCalculatedItem = 572,
        yxDialogPivotClientServerSet = 689,
        yxDialogPivotFieldGroup = 433,
        yxDialogPivotFieldProperties = 313,
        yxDialogPivotFieldUngroup = 434,
        yxDialogPivotShowPages = 421,
        yxDialogPivotSolveOrder = 568,
        yxDialogPivotTableOptions = 567,
        yxDialogPivotTableWizard = 312,
        yxDialogPlacement = 300,
        yxDialogPrint = 8,
        yxDialogPrinterSetup = 9,
        yxDialogPrintPreview = 222,
        yxDialogPromote = 202,
        yxDialogProperties = 474,
        yxDialogPropertyFields = 754,
        yxDialogProtectDocument = 28,
        yxDialogProtectSharing = 620,
        yxDialogPublishAsWebPage = 653,
        yxDialogPushbuttonProperties = 445,
        yxDialogReplaceFont = 134,
        yxDialogRoutingSlip = 336,
        yxDialogRowHeight = 127,
        yxDialogRun = 17,
        yxDialogSaveAs = 5,
        yxDialogSaveCopyAs = 456,
        yxDialogSaveNewObject = 208,
        yxDialogSaveWorkbook = 145,
        yxDialogSaveWorkspace = 285,
        yxDialogScale = 87,
        yxDialogScenarioAdd = 307,
        yxDialogScenarioCells = 305,
        yxDialogScenarioEdit = 308,
        yxDialogScenarioMerge = 473,
        yxDialogScenarioSummary = 311,
        yxDialogScrollbarProperties = 420,
        yxDialogSearch = 731,
        yxDialogSelectSpecial = 132,
        yxDialogSendMail = 189,
        yxDialogSeriesAxes = 460,
        yxDialogSeriesOptions = 557,
        yxDialogSeriesOrder = 466,
        yxDialogSeriesShape = 504,
        yxDialogSeriesX = 461,
        yxDialogSeriesY = 462,
        yxDialogSetBackgroundPicture = 509,
        yxDialogSetPrintTitles = 23,
        yxDialogSetUpdateStatus = 159,
        yxDialogShowDetail = 204,
        yxDialogShowToolbar = 220,
        yxDialogSize = 261,
        yxDialogSort = 39,
        yxDialogSortSpecial = 192,
        yxDialogSplit = 137,
        yxDialogStandardFont = 190,
        yxDialogStandardWidth = 472,
        yxDialogStyle = 44,
        yxDialogSubscribeTo = 218,
        yxDialogSubtotalCreate = 398,
        yxDialogSummaryInfo = 474,
        yxDialogTable = 41,
        yxDialogTabOrder = 394,
        yxDialogTextToColumns = 422,
        yxDialogUnhide = 94,
        yxDialogUpdateLink = 201,
        yxDialogVbaInsertFile = 328,
        yxDialogVbaMakeAddin = 478,
        yxDialogVbaProcedureDefinition = 330,
        yxDialogView3d = 197,
        yxDialogWebOptionsBrowsers = 773,
        yxDialogWebOptionsEncoding = 686,
        yxDialogWebOptionsFiles = 684,
        yxDialogWebOptionsFonts = 687,
        yxDialogWebOptionsGeneral = 683,
        yxDialogWebOptionsPictures = 685,
        yxDialogWindowMove = 14,
        yxDialogWindowSize = 13,
        yxDialogWorkbookAdd = 281,
        yxDialogWorkbookCopy = 283,
        yxDialogWorkbookInsert = 354,
        yxDialogWorkbookMove = 282,
        yxDialogWorkbookName = 386,
        yxDialogWorkbookNew = 302,
        yxDialogWorkbookOptions = 284,
        yxDialogWorkbookProtect = 417,
        yxDialogWorkbookTabSplit = 415,
        yxDialogWorkbookUnhide = 384,
        yxDialogWorkgroup = 199,
        yxDialogWorkspace = 95,
        yxDialogZoom = 256

    }
    export const enum YxCalculatedMemberType {
        yxCalculatedMember = 0,
        yxCalculatedSet = 1

    }
    export const enum YxCalculatedMemberType {
        yxCalculatedMember = 0,
        yxCalculatedSet = 1

    }
    export const enum YxCalculation {
        yxCalculationAutomatic = -4105,
        yxCalculationManual = -4135,
        yxCalculationSemiautomatic = 2

    }
    export const enum YxCalculation {
        yxCalculationAutomatic = -4105,
        yxCalculationManual = -4135,
        yxCalculationSemiautomatic = 2

    }
    export const enum YxCalculationInterruptKey {
        yxAnyKey = 2,
        yxEscKey = 1,
        yxNoKey = 0

    }
    export const enum YxCalculationInterruptKey {
        yxAnyKey = 2,
        yxEscKey = 1,
        yxNoKey = 0

    }
    export const enum YxCalculationState {
        yxCalculating = 1,
        yxDone = 0,
        yxPending = 2

    }
    export const enum YxCalculationState {
        yxCalculating = 1,
        yxDone = 0,
        yxPending = 2

    }
    export const enum YxCategoryType {
        yxAutomaticScale = -4105,
        yxCategoryScale = 2,
        yxTimeScale = 3

    }
    export const enum YxCategoryType {
        yxAutomaticScale = -4105,
        yxCategoryScale = 2,
        yxTimeScale = 3

    }
    export const enum YxCellType {
        yxCellTypeAllFormatConditions = -4172,
        yxCellTypeAllValidation = -4174,
        yxCellTypeBlanks = 4,
        yxCellTypeComments = -4144,
        yxCellTypeConstants = 2,
        yxCellTypeFormulas = -4123,
        yxCellTypeLastCell = 11,
        yxCellTypeSameFormatConditions = -4173,
        yxCellTypeSameValidation = -4175,
        yxCellTypeVisible = 12

    }
    export const enum YxCellType {
        yxCellTypeAllFormatConditions = -4172,
        yxCellTypeAllValidation = -4174,
        yxCellTypeBlanks = 4,
        yxCellTypeComments = -4144,
        yxCellTypeConstants = 2,
        yxCellTypeFormulas = -4123,
        yxCellTypeLastCell = 11,
        yxCellTypeSameFormatConditions = -4173,
        yxCellTypeSameValidation = -4175,
        yxCellTypeVisible = 12

    }
    export const enum YxChartLocation {
        yxLocationAsNewSheet = 1,
        yxLocationAsObject = 2,
        yxLocationAutomatic = 3

    }
    export const enum YxChartLocation {
        yxLocationAsNewSheet = 1,
        yxLocationAsObject = 2,
        yxLocationAutomatic = 3

    }
    export const enum YxChartPicturePlacement {
        yxAllFaces = 7,
        yxEnd = 2,
        yxEndSides = 3,
        yxFront = 4,
        yxFrontEnd = 6,
        yxFrontSides = 5,
        yxSides = 1

    }
    export const enum YxChartPicturePlacement {
        yxAllFaces = 7,
        yxEnd = 2,
        yxEndSides = 3,
        yxFront = 4,
        yxFrontEnd = 6,
        yxFrontSides = 5,
        yxSides = 1

    }
    export const enum YxChartPictureType {
        yxStack = 2,
        yxStackScale = 3,
        yxStretch = 1

    }
    export const enum YxChartPictureType {
        yxStack = 2,
        yxStackScale = 3,
        yxStretch = 1

    }
    export const enum YxChartSplitType {
        yxSplitByCustomSplit = 4,
        yxSplitByPercentValue = 3,
        yxSplitByPosition = 1,
        yxSplitByValue = 2

    }
    export const enum YxChartSplitType {
        yxSplitByCustomSplit = 4,
        yxSplitByPercentValue = 3,
        yxSplitByPosition = 1,
        yxSplitByValue = 2

    }
    export const enum YxChartType {
        yxArea = 1,
        yxAreaStacked = 76,
        yxAreaStacked100 = 77,
        yx3DArea = -4098,
        yx3DAreaStacked = 78,
        yx3DAreaStacked100 = 79,
        yxBarClustered = 57,
        yxBarOfPie = 71,
        yxBarStacked = 58,
        yxBarStacked100 = 59,
        yx3DBarClustered = 60,
        yx3DBarStacked = 61,
        yx3DBarStacked100 = 62,
        yxBubble = 15,
        yxBubble3DEffect = 87,
        yxColumnClustered = 51,
        yxColumnStacked = 52,
        yxColumnStacked100 = 53,
        yx3DColumn = -4100,
        yx3DColumnClustered = 54,
        yx3DColumnStacked = 55,
        yx3DColumnStacked100 = 56,
        yxConeBarClustered = 102,
        yxConeBarStacked = 103,
        yxConeBarStacked100 = 104,
        yxConeCol = 105,
        yxConeColClustered = 99,
        yxConeColStacked = 100,
        yxConeColStacked100 = 101,
        yxCylinderBarClustered = 95,
        yxCylinderBarStacked = 96,
        yxCylinderBarStacked100 = 97,
        yxCylinderCol = 98,
        yxCylinderColClustered = 92,
        yxCylinderColStacked = 93,
        yxCylinderColStacked100 = 94,
        yxDoughnut = -4120,
        yxDoughnutExploded = 80,
        yxLine = 4,
        yxLineMarkers = 65,
        yxLineMarkersStacked = 66,
        yxLineMarkersStacked100 = 67,
        yxLineStacked = 63,
        yxLineStacked100 = 64,
        yx3DLine = -4101,
        yxPie = 5,
        yxPieExploded = 69,
        yxPieOfPie = 68,
        yx3DPie = -4102,
        yx3DPieExploded = 70,
        yxPyramidBarClustered = 109,
        yxPyramidBarStacked = 110,
        yxPyramidBarStacked100 = 111,
        yxPyramidCol = 112,
        yxPyramidColClustered = 106,
        yxPyramidColStacked = 107,
        yxPyramidColStacked100 = 108,
        yxRadar = -4151,
        yxRadarFilled = 82,
        yxRadarMarkers = 81,
        yxStockHLC = 88,
        yxStockOHLC = 89,
        yxStockVHLC = 90,
        yxStockVOHLC = 91,
        yxSurface = 83,
        yxSurfaceTopView = 85,
        yxSurfaceTopViewWireframe = 86,
        yxSurfaceWireframe = 84,
        yxXYScatter = -4169,
        yxXYScatterLines = 74,
        yxXYScatterLinesNoMarkers = 75,
        yxXYScatterSmooth = 72,
        yxXYScatterSmoothNoMarkers = 73

    }
    export const enum YxChartType {
        yxArea = 1,
        yxAreaStacked = 76,
        yxAreaStacked100 = 77,
        yx3DArea = -4098,
        yx3DAreaStacked = 78,
        yx3DAreaStacked100 = 79,
        yxBarClustered = 57,
        yxBarOfPie = 71,
        yxBarStacked = 58,
        yxBarStacked100 = 59,
        yx3DBarClustered = 60,
        yx3DBarStacked = 61,
        yx3DBarStacked100 = 62,
        yxBubble = 15,
        yxBubble3DEffect = 87,
        yxColumnClustered = 51,
        yxColumnStacked = 52,
        yxColumnStacked100 = 53,
        yx3DColumn = -4100,
        yx3DColumnClustered = 54,
        yx3DColumnStacked = 55,
        yx3DColumnStacked100 = 56,
        yxConeBarClustered = 102,
        yxConeBarStacked = 103,
        yxConeBarStacked100 = 104,
        yxConeCol = 105,
        yxConeColClustered = 99,
        yxConeColStacked = 100,
        yxConeColStacked100 = 101,
        yxCylinderBarClustered = 95,
        yxCylinderBarStacked = 96,
        yxCylinderBarStacked100 = 97,
        yxCylinderCol = 98,
        yxCylinderColClustered = 92,
        yxCylinderColStacked = 93,
        yxCylinderColStacked100 = 94,
        yxDoughnut = -4120,
        yxDoughnutExploded = 80,
        yxLine = 4,
        yxLineMarkers = 65,
        yxLineMarkersStacked = 66,
        yxLineMarkersStacked100 = 67,
        yxLineStacked = 63,
        yxLineStacked100 = 64,
        yx3DLine = -4101,
        yxPie = 5,
        yxPieExploded = 69,
        yxPieOfPie = 68,
        yx3DPie = -4102,
        yx3DPieExploded = 70,
        yxPyramidBarClustered = 109,
        yxPyramidBarStacked = 110,
        yxPyramidBarStacked100 = 111,
        yxPyramidCol = 112,
        yxPyramidColClustered = 106,
        yxPyramidColStacked = 107,
        yxPyramidColStacked100 = 108,
        yxRadar = -4151,
        yxRadarFilled = 82,
        yxRadarMarkers = 81,
        yxStockHLC = 88,
        yxStockOHLC = 89,
        yxStockVHLC = 90,
        yxStockVOHLC = 91,
        yxSurface = 83,
        yxSurfaceTopView = 85,
        yxSurfaceTopViewWireframe = 86,
        yxSurfaceWireframe = 84,
        yxXYScatter = -4169,
        yxXYScatterLines = 74,
        yxXYScatterLinesNoMarkers = 75,
        yxXYScatterSmooth = 72,
        yxXYScatterSmoothNoMarkers = 73

    }
    export const enum YxColorIndex {
        yxColorIndexAutomatic = -4105,
        yxColorIndexNone = -4142

    }
    export const enum YxColorIndex {
        yxColorIndexAutomatic = -4105,
        yxColorIndexNone = -4142

    }
    export const enum YxColumnDataType {
        yxDMYFormat = 4,
        yxDYMFormat = 7,
        yxEMDFormat = 10,
        yxGeneralFormat = 1,
        yxMDYFormat = 3,
        yxMYDFormat = 6,
        yxSkipColumn = 9,
        yxTextFormat = 2,
        yxYDMFormat = 8,
        yxYMDFormat = 5

    }
    export const enum YxColumnDataType {
        yxDMYFormat = 4,
        yxDYMFormat = 7,
        yxEMDFormat = 10,
        yxGeneralFormat = 1,
        yxMDYFormat = 3,
        yxMYDFormat = 6,
        yxSkipColumn = 9,
        yxTextFormat = 2,
        yxYDMFormat = 8,
        yxYMDFormat = 5

    }
    export const enum YxCommentDisplayMode {
        yxCommentAndIndicator = 1,
        yxCommentIndicatorOnly = -1,
        yxNoIndicator = 0

    }
    export const enum YxCommentDisplayMode {
        yxCommentAndIndicator = 1,
        yxCommentIndicatorOnly = -1,
        yxNoIndicator = 0

    }
    export const enum YxConditionValueTypes {
        xlConditionValueAutomaticMax = 7,
        xlConditionValueAutomaticMin = 6,
        xlConditionValueFormula = 4,
        xlConditionValueHighestValue = 2,
        xlConditionValueLowestValue = 1,
        xlConditionValueNone = -1,
        xlConditionValueNumber = 0,
        xlConditionValuePercent = 3,
        xlConditionValuePercentile = 5

    }
    export const enum YxConditionValueTypes {
        xlConditionValueAutomaticMax = 7,
        xlConditionValueAutomaticMin = 6,
        xlConditionValueFormula = 4,
        xlConditionValueHighestValue = 2,
        xlConditionValueLowestValue = 1,
        xlConditionValueNone = -1,
        xlConditionValueNumber = 0,
        xlConditionValuePercent = 3,
        xlConditionValuePercentile = 5

    }
    export const enum YxConsolidationFunction {
        yxAverage = -4106,
        yxCount = -4112,
        yxCountNums = -4113,
        yxMax = -4136,
        yxMin = -4139,
        yxProduct = -4149,
        yxStDev = -4155,
        yxStDevP = -4156,
        yxSum = -4157,
        yxUnknown = 1000,
        yxVar = -4164,
        yxVarP = -4165

    }
    export const enum YxConsolidationFunction {
        yxAverage = -4106,
        yxCount = -4112,
        yxCountNums = -4113,
        yxMax = -4136,
        yxMin = -4139,
        yxProduct = -4149,
        yxStDev = -4155,
        yxStDevP = -4156,
        yxSum = -4157,
        yxUnknown = 1000,
        yxVar = -4164,
        yxVarP = -4165

    }
    export const enum YxConstants {
        yx3DBar = -4099,
        yx3DEffects1 = 13,
        yx3DEffects2 = 14,
        yx3DSurface = -4103,
        yxAbove = 0,
        yxAccounting1 = 4,
        yxAccounting2 = 5,
        yxAccounting4 = 17,
        yxAdd = 2,
        yxAll = -4104,
        yxAccounting3 = 6,
        yxAllExceptBorders = 7,
        yxAutomatic = -4105,
        yxBar = 2,
        yxBelow = 1,
        yxBidi = -5000,
        yxBidiCalendar = 3,
        yxBoth = 1,
        yxBottom = -4107,
        yxCascade = 7,
        yxCenter = -4108,
        yxCenterAcrossSelection = 7,
        yxChart4 = 2,
        yxChartSeries = 17,
        yxChartShort = 6,
        yxChartTitles = 18,
        yxChecker = 9,
        yxCircle = 8,
        yxClassic1 = 1,
        yxClassic2 = 2,
        yxClassic3 = 3,
        yxClosed = 3,
        yxColor1 = 7,
        yxColor2 = 8,
        yxColor3 = 9,
        yxColumn = 3,
        yxCombination = -4111,
        yxComplete = 4,
        yxConstants = 2,
        yxContents = 2,
        yxContext = -5002,
        yxCorner = 2,
        yxCrissCross = 16,
        yxCross = 4,
        yxCustom = -4114,
        yxDebugCodePane = 13,
        yxDefaultAutoFormat = -1,
        yxDesktop = 9,
        yxDiamond = 2,
        yxDirect = 1,
        yxDistributed = -4117,
        yxDivide = 5,
        yxDoubleAccounting = 5,
        yxDoubleClosed = 5,
        yxDoubleOpen = 4,
        yxDoubleQuote = 1,
        yxDrawingObject = 14,
        yxEntireChart = 20,
        yxExcelMenus = 1,
        yxExtended = 3,
        yxFill = 5,
        yxFirst = 0,
        yxFixedValue = 1,
        yxFloating = 5,
        yxFormats = -4122,
        yxFormula = 5,
        yxFullScript = 1,
        yxGeneral = 1,
        yxGray16 = 17,
        yxGray25 = -4124,
        yxGray50 = -4125,
        yxGray75 = -4126,
        yxGray8 = 18,
        yxGregorian = 2,
        yxGrid = 15,
        yxGridline = 22,
        yxHigh = -4127,
        yxHindiNumerals = 3,
        yxIcons = 1,
        yxImmediatePane = 12,
        yxInside = 2,
        yxInteger = 2,
        yxJustify = -4130,
        yxLast = 1,
        yxLastCell = 11,
        yxLatin = -5001,
        yxLeft = -4131,
        yxLeftToRight = 2,
        yxLightDown = 13,
        yxLightHorizontal = 11,
        yxLightUp = 14,
        yxLightVertical = 12,
        yxList1 = 10,
        yxList2 = 11,
        yxList3 = 12,
        yxLocalFormat1 = 15,
        yxLocalFormat2 = 16,
        yxLogicalCursor = 1,
        yxLong = 3,
        yxLotusHelp = 2,
        yxLow = -4134,
        yxLTR = -5003,
        yxMacrosheetCell = 7,
        yxManual = -4135,
        yxMaximum = 2,
        yxMinimum = 4,
        yxMinusValues = 3,
        yxMixed = 2,
        yxMixedAuthorizedScript = 4,
        yxMixedScript = 3,
        yxModule = -4141,
        yxMultiply = 4,
        yxNarrow = 1,
        yxNextToAxis = 4,
        yxNoDocuments = 3,
        yxNone = -4142,
        yxNotes = -4144,
        yxOff = -4146,
        yxOn = 1,
        yxOpaque = 3,
        yxOpen = 2,
        yxOutside = 3,
        yxPartial = 3,
        yxPartialScript = 2,
        yxPercent = 2,
        yxPlus = 9,
        yxPlusValues = 2,
        yxReference = 4,
        yxRight = -4152,
        yxRTL = -5004,
        yxScale = 3,
        yxSemiautomatic = 2,
        yxSemiGray75 = 10,
        yxShort = 1,
        yxShowLabel = 4,
        yxShowLabelAndPercent = 5,
        yxShowPercent = 3,
        yxShowValue = 2,
        yxSimple = -4154,
        yxSingle = 2,
        yxSingleAccounting = 4,
        yxSingleQuote = 2,
        yxSolid = 1,
        yxSquare = 1,
        yxStar = 5,
        yxStError = 4,
        yxStrict = 2,
        yxSubtract = 3,
        yxSystem = 1,
        yxTextBox = 16,
        yxTiled = 1,
        yxTitleBar = 8,
        yxToolbar = 1,
        yxToolbarButton = 2,
        yxTop = -4160,
        yxTopToBottom = 1,
        yxTransparent = 2,
        yxTriangle = 3,
        yxVeryHidden = 2,
        yxVisible = 12,
        yxVisualCursor = 2,
        yxWatchPane = 11,
        yxWide = 3,
        yxWorkbookTab = 6,
        yxWorksheet4 = 1,
        yxWorksheetCell = 3,
        yxWorksheetShort = 5

    }
    export const enum YxConstants {
        yx3DBar = -4099,
        yx3DEffects1 = 13,
        yx3DEffects2 = 14,
        yx3DSurface = -4103,
        yxAbove = 0,
        yxAccounting1 = 4,
        yxAccounting2 = 5,
        yxAccounting4 = 17,
        yxAdd = 2,
        yxAll = -4104,
        yxAccounting3 = 6,
        yxAllExceptBorders = 7,
        yxAutomatic = -4105,
        yxBar = 2,
        yxBelow = 1,
        yxBidi = -5000,
        yxBidiCalendar = 3,
        yxBoth = 1,
        yxBottom = -4107,
        yxCascade = 7,
        yxCenter = -4108,
        yxCenterAcrossSelection = 7,
        yxChart4 = 2,
        yxChartSeries = 17,
        yxChartShort = 6,
        yxChartTitles = 18,
        yxChecker = 9,
        yxCircle = 8,
        yxClassic1 = 1,
        yxClassic2 = 2,
        yxClassic3 = 3,
        yxClosed = 3,
        yxColor1 = 7,
        yxColor2 = 8,
        yxColor3 = 9,
        yxColumn = 3,
        yxCombination = -4111,
        yxComplete = 4,
        yxConstants = 2,
        yxContents = 2,
        yxContext = -5002,
        yxCorner = 2,
        yxCrissCross = 16,
        yxCross = 4,
        yxCustom = -4114,
        yxDebugCodePane = 13,
        yxDefaultAutoFormat = -1,
        yxDesktop = 9,
        yxDiamond = 2,
        yxDirect = 1,
        yxDistributed = -4117,
        yxDivide = 5,
        yxDoubleAccounting = 5,
        yxDoubleClosed = 5,
        yxDoubleOpen = 4,
        yxDoubleQuote = 1,
        yxDrawingObject = 14,
        yxEntireChart = 20,
        yxExcelMenus = 1,
        yxExtended = 3,
        yxFill = 5,
        yxFirst = 0,
        yxFixedValue = 1,
        yxFloating = 5,
        yxFormats = -4122,
        yxFormula = 5,
        yxFullScript = 1,
        yxGeneral = 1,
        yxGray16 = 17,
        yxGray25 = -4124,
        yxGray50 = -4125,
        yxGray75 = -4126,
        yxGray8 = 18,
        yxGregorian = 2,
        yxGrid = 15,
        yxGridline = 22,
        yxHigh = -4127,
        yxHindiNumerals = 3,
        yxIcons = 1,
        yxImmediatePane = 12,
        yxInside = 2,
        yxInteger = 2,
        yxJustify = -4130,
        yxLast = 1,
        yxLastCell = 11,
        yxLatin = -5001,
        yxLeft = -4131,
        yxLeftToRight = 2,
        yxLightDown = 13,
        yxLightHorizontal = 11,
        yxLightUp = 14,
        yxLightVertical = 12,
        yxList1 = 10,
        yxList2 = 11,
        yxList3 = 12,
        yxLocalFormat1 = 15,
        yxLocalFormat2 = 16,
        yxLogicalCursor = 1,
        yxLong = 3,
        yxLotusHelp = 2,
        yxLow = -4134,
        yxLTR = -5003,
        yxMacrosheetCell = 7,
        yxManual = -4135,
        yxMaximum = 2,
        yxMinimum = 4,
        yxMinusValues = 3,
        yxMixed = 2,
        yxMixedAuthorizedScript = 4,
        yxMixedScript = 3,
        yxModule = -4141,
        yxMultiply = 4,
        yxNarrow = 1,
        yxNextToAxis = 4,
        yxNoDocuments = 3,
        yxNone = -4142,
        yxNotes = -4144,
        yxOff = -4146,
        yxOn = 1,
        yxOpaque = 3,
        yxOpen = 2,
        yxOutside = 3,
        yxPartial = 3,
        yxPartialScript = 2,
        yxPercent = 2,
        yxPlus = 9,
        yxPlusValues = 2,
        yxReference = 4,
        yxRight = -4152,
        yxRTL = -5004,
        yxScale = 3,
        yxSemiautomatic = 2,
        yxSemiGray75 = 10,
        yxShort = 1,
        yxShowLabel = 4,
        yxShowLabelAndPercent = 5,
        yxShowPercent = 3,
        yxShowValue = 2,
        yxSimple = -4154,
        yxSingle = 2,
        yxSingleAccounting = 4,
        yxSingleQuote = 2,
        yxSolid = 1,
        yxSquare = 1,
        yxStar = 5,
        yxStError = 4,
        yxStrict = 2,
        yxSubtract = 3,
        yxSystem = 1,
        yxTextBox = 16,
        yxTiled = 1,
        yxTitleBar = 8,
        yxToolbar = 1,
        yxToolbarButton = 2,
        yxTop = -4160,
        yxTopToBottom = 1,
        yxTransparent = 2,
        yxTriangle = 3,
        yxVeryHidden = 2,
        yxVisible = 12,
        yxVisualCursor = 2,
        yxWatchPane = 11,
        yxWide = 3,
        yxWorkbookTab = 6,
        yxWorksheet4 = 1,
        yxWorksheetCell = 3,
        yxWorksheetShort = 5

    }
    export const enum YxCopyPictureFormat {
        yxBitmap = 2,
        yxPicture = -4147

    }
    export const enum YxCopyPictureFormat {
        yxBitmap = 2,
        yxPicture = -4147

    }
    export const enum YxCreator {
        yxCreatorCode = 1480803660

    }
    export const enum YxCreator {
        yxCreatorCode = 1480803660

    }
    export const enum YxCubeFieldType {
        yxHierarchy = 1,
        yxMeasure = 2,
        yxSet = 3

    }
    export const enum YxCubeFieldType {
        yxHierarchy = 1,
        yxMeasure = 2,
        yxSet = 3

    }
    export const enum YxCutCopyMode {
        yxCopy = 1,
        yxCut = 2

    }
    export const enum YxCutCopyMode {
        yxCopy = 1,
        yxCut = 2

    }
    export const enum YxDataBarAxisPosition {
        xlDataBarAxisAutomatic = 0,
        xlDataBarAxisMidpoint = 1,
        xlDataBarAxisNone = 2

    }
    export const enum YxDataBarAxisPosition {
        xlDataBarAxisAutomatic = 0,
        xlDataBarAxisMidpoint = 1,
        xlDataBarAxisNone = 2

    }
    export const enum YxDataBarBorderType {
        xlDataBarBorderNone = 0,
        xlDataBarBorderSolid = 1

    }
    export const enum YxDataBarBorderType {
        xlDataBarBorderNone = 0,
        xlDataBarBorderSolid = 1

    }
    export const enum YxDataBarFillType {
        xlDataBarFillGradient = 1,
        xlDataBarFillSolid = 0

    }
    export const enum YxDataBarFillType {
        xlDataBarFillGradient = 1,
        xlDataBarFillSolid = 0

    }
    export const enum YxDataBarNegativeColorType {
        xlDataBarColor = 0,
        xlDataBarSameAsPositive = 1

    }
    export const enum YxDataBarNegativeColorType {
        xlDataBarColor = 0,
        xlDataBarSameAsPositive = 1

    }
    export const enum YxDataLabelsType {
        yxDataLabelsShowBubbleSizes = 6,
        yxDataLabelsShowLabel = 4,
        yxDataLabelsShowLabelAndPercent = 5,
        yxDataLabelsShowNone = -4142,
        yxDataLabelsShowPercent = 3,
        yxDataLabelsShowValue = 2

    }
    export const enum YxDataLabelsType {
        yxDataLabelsShowBubbleSizes = 6,
        yxDataLabelsShowLabel = 4,
        yxDataLabelsShowLabelAndPercent = 5,
        yxDataLabelsShowNone = -4142,
        yxDataLabelsShowPercent = 3,
        yxDataLabelsShowValue = 2

    }
    export const enum YxDataSeriesDate {
        yxDay = 1,
        yxMonth = 3,
        yxWeekday = 2,
        yxYear = 4

    }
    export const enum YxDataSeriesDate {
        yxDay = 1,
        yxMonth = 3,
        yxWeekday = 2,
        yxYear = 4

    }
    export const enum YxDataSeriesType {
        yxAutoFill = 4,
        yxChronological = 3,
        yxDataSeriesLinear = -4132,
        yxGrowth = 2

    }
    export const enum YxDataSeriesType {
        yxAutoFill = 4,
        yxChronological = 3,
        yxDataSeriesLinear = -4132,
        yxGrowth = 2

    }
    export const enum YxDeleteShiftDirection {
        yxShiftToLeft = -4159,
        yxShiftUp = -4162

    }
    export const enum YxDeleteShiftDirection {
        yxShiftToLeft = -4159,
        yxShiftUp = -4162

    }
    export const enum YxDirection {
        yxDown = -4121,
        yxToLeft = -4159,
        yxToRight = -4161,
        yxUp = -4162

    }
    export const enum YxDirection {
        yxDown = -4121,
        yxToLeft = -4159,
        yxToRight = -4161,
        yxUp = -4162

    }
    export const enum YxDisplayBlanksAs {
        yxInterpolated = 3,
        yxNotPlotted = 1,
        yxZero = 2

    }
    export const enum YxDisplayBlanksAs {
        yxInterpolated = 3,
        yxNotPlotted = 1,
        yxZero = 2

    }
    export const enum YxDisplayDrawingObjects {
        yxDisplayShapes = -4104,
        yxHide = 3,
        yxPlaceholders = 2

    }
    export const enum YxDisplayDrawingObjects {
        yxDisplayShapes = -4104,
        yxHide = 3,
        yxPlaceholders = 2

    }
    export const enum YxDisplayUnit {
        yxHundredMillions = -8,
        yxHundreds = -2,
        yxHundredThousands = -5,
        yxMillionMillions = -10,
        yxMillions = -6,
        yxTenMillions = -7,
        yxTenThousands = -4,
        yxThousandMillions = -9,
        yxThousands = -3

    }
    export const enum YxDisplayUnit {
        yxHundredMillions = -8,
        yxHundreds = -2,
        yxHundredThousands = -5,
        yxMillionMillions = -10,
        yxMillions = -6,
        yxTenMillions = -7,
        yxTenThousands = -4,
        yxThousandMillions = -9,
        yxThousands = -3

    }
    export const enum YxDupeUnique {
        xlDuplicate = 1,
        xlUnique = 0

    }
    export const enum YxDupeUnique {
        xlDuplicate = 1,
        xlUnique = 0

    }
    export const enum YxDVAlertStyle {
        yxValidAlertInformation = 3,
        yxValidAlertStop = 1,
        yxValidAlertWarning = 2

    }
    export const enum YxDVAlertStyle {
        yxValidAlertInformation = 3,
        yxValidAlertStop = 1,
        yxValidAlertWarning = 2

    }
    export const enum YxDVType {
        yxValidateCustom = 7,
        yxValidateDate = 4,
        yxValidateDecimal = 2,
        yxValidateInputOnly = 0,
        yxValidateList = 3,
        yxValidateTextLength = 6,
        yxValidateTime = 5,
        yxValidateWholeNumber = 1

    }
    export const enum YxDVType {
        yxValidateCustom = 7,
        yxValidateDate = 4,
        yxValidateDecimal = 2,
        yxValidateInputOnly = 0,
        yxValidateList = 3,
        yxValidateTextLength = 6,
        yxValidateTime = 5,
        yxValidateWholeNumber = 1

    }
    export const enum YxEditionOptionsOption {
        yxAutomaticUpdate = 4,
        yxCancel = 1,
        yxChangeAttributes = 6,
        yxManualUpdate = 5,
        yxOpenSource = 3,
        yxSelect = 3,
        yxSendPublisher = 2,
        yxUpdateSubscriber = 2

    }
    export const enum YxEditionOptionsOption {
        yxAutomaticUpdate = 4,
        yxCancel = 1,
        yxChangeAttributes = 6,
        yxManualUpdate = 5,
        yxOpenSource = 3,
        yxSelect = 3,
        yxSendPublisher = 2,
        yxUpdateSubscriber = 2

    }
    export const enum YxEditionType {
        yxPublisher = 1,
        yxSubscriber = 2

    }
    export const enum YxEditionType {
        yxPublisher = 1,
        yxSubscriber = 2

    }
    export const enum YxEnableSelection {
        yxNoRestrictions = 0,
        yxNoSelection = -4142,
        yxUnlockedCells = 1

    }
    export const enum YxEnableSelection {
        yxNoRestrictions = 0,
        yxNoSelection = -4142,
        yxUnlockedCells = 1

    }
    export const enum YxErrorBarDirection {
        yxX = -4168,
        yxY = 1

    }
    export const enum YxErrorBarDirection {
        yxX = -4168,
        yxY = 1

    }
    export const enum YxErrorBarInclude {
        yxErrorBarIncludeBoth = 1,
        yxErrorBarIncludeMinusValues = 3,
        yxErrorBarIncludeNone = -4142,
        yxErrorBarIncludePlusValues = 2

    }
    export const enum YxErrorBarInclude {
        yxErrorBarIncludeBoth = 1,
        yxErrorBarIncludeMinusValues = 3,
        yxErrorBarIncludeNone = -4142,
        yxErrorBarIncludePlusValues = 2

    }
    export const enum YxErrorBarType {
        yxErrorBarTypeCustom = -4114,
        yxErrorBarTypeFixedValue = 1,
        yxErrorBarTypePercent = 2,
        yxErrorBarTypeStDev = -4155,
        yxErrorBarTypeStError = 4

    }
    export const enum YxErrorBarType {
        yxErrorBarTypeCustom = -4114,
        yxErrorBarTypeFixedValue = 1,
        yxErrorBarTypePercent = 2,
        yxErrorBarTypeStDev = -4155,
        yxErrorBarTypeStError = 4

    }
    export const enum YxFileAccess {
        yxReadOnly = 3,
        yxReadWrite = 2

    }
    export const enum YxFileAccess {
        yxReadOnly = 3,
        yxReadWrite = 2

    }
    export const enum YxFileFormat {
        yxAddIn = 18,
        xlAddIn8 = 18,
        yxCSV = 6,
        yxCSVMac = 22,
        yxCSVMSDOS = 24,
        yxCSVWindows = 23,
        yxCurrentPlatformText = -4158,
        yxDBF2 = 7,
        yxDBF3 = 8,
        yxDBF4 = 11,
        yxDIF = 9,
        yxExcel2 = 16,
        yxExcel2FarEast = 27,
        yxExcel3 = 29,
        yxExcel4 = 33,
        yxExcel4Workbook = 35,
        yxExcel5 = 39,
        yxExcel7 = 39,
        yxExcel9795 = 43,
        yxHtml = 44,
        yxIntlAddIn = 26,
        yxIntlMacro = 25,
        yxOpenDocumentSpreadsheet = 60,
        yxOpenXMLAddIn = 55,
        yxOpenXMLTemplate = 54,
        yxOpenXMLTemplateMacroEnabled = 53,
        yxOpenXMLWorkbook = 51,
        yxOpenXMLWorkbookMacroEnabled = 52,
        yxSYLK = 2,
        yxTemplate = 17,
        yxTextMac = 19,
        yxTextMSDOS = 21,
        yxTextPrinter = 36,
        yxTextWindows = 20,
        yxUnicodeText = 42,
        yxWebArchive = 45,
        yxWJ2WD1 = 14,
        yxWJ3 = 40,
        yxWJ3FJ3 = 41,
        yxWK1 = 5,
        yxWK1ALL = 31,
        yxWK1FMT = 30,
        yxWK3 = 15,
        yxWK3FM3 = 32,
        yxWK4 = 38,
        yxWKS = 4,
        yxWorkbookNormal = -4143,
        yxWorks2FarEast = 28,
        yxWQ1 = 34,
        yxXMLSpreadsheet = 46,
        yxWorkbookDefault = 51

    }
    export const enum YxFileFormat {
        yxAddIn = 18,
        xlAddIn8 = 18,
        yxCSV = 6,
        yxCSVMac = 22,
        yxCSVMSDOS = 24,
        yxCSVWindows = 23,
        yxCurrentPlatformText = -4158,
        yxDBF2 = 7,
        yxDBF3 = 8,
        yxDBF4 = 11,
        yxDIF = 9,
        yxExcel2 = 16,
        yxExcel2FarEast = 27,
        yxExcel3 = 29,
        yxExcel4 = 33,
        yxExcel4Workbook = 35,
        yxExcel5 = 39,
        yxExcel7 = 39,
        yxExcel9795 = 43,
        yxHtml = 44,
        yxIntlAddIn = 26,
        yxIntlMacro = 25,
        yxOpenDocumentSpreadsheet = 60,
        yxOpenXMLAddIn = 55,
        yxOpenXMLTemplate = 54,
        yxOpenXMLTemplateMacroEnabled = 53,
        yxOpenXMLWorkbook = 51,
        yxOpenXMLWorkbookMacroEnabled = 52,
        yxSYLK = 2,
        yxTemplate = 17,
        yxTextMac = 19,
        yxTextMSDOS = 21,
        yxTextPrinter = 36,
        yxTextWindows = 20,
        yxUnicodeText = 42,
        yxWebArchive = 45,
        yxWJ2WD1 = 14,
        yxWJ3 = 40,
        yxWJ3FJ3 = 41,
        yxWK1 = 5,
        yxWK1ALL = 31,
        yxWK1FMT = 30,
        yxWK3 = 15,
        yxWK3FM3 = 32,
        yxWK4 = 38,
        yxWKS = 4,
        yxWorkbookNormal = -4143,
        yxWorks2FarEast = 28,
        yxWQ1 = 34,
        yxXMLSpreadsheet = 46,
        yxWorkbookDefault = 51

    }
    export const enum YxFillWith {
        yxFillWithAll = -4104,
        yxFillWithContents = 2,
        yxFillWithFormats = -4122

    }
    export const enum YxFillWith {
        yxFillWithAll = -4104,
        yxFillWithContents = 2,
        yxFillWithFormats = -4122

    }
    export const enum YxFilterAction {
        yxFilterCopy = 2,
        yxFilterInPlace = 1

    }
    export const enum YxFilterAction {
        yxFilterCopy = 2,
        yxFilterInPlace = 1

    }
    export const enum YxFormatConditionOperator {
        yxBetween = 1,
        yxEqual = 3,
        yxGreater = 5,
        yxGreaterEqual = 7,
        yxLess = 6,
        yxLessEqual = 8,
        yxNotBetween = 2,
        yxNotEqual = 4

    }
    export const enum YxFormatConditionOperator {
        yxBetween = 1,
        yxEqual = 3,
        yxGreater = 5,
        yxGreaterEqual = 7,
        yxLess = 6,
        yxLessEqual = 8,
        yxNotBetween = 2,
        yxNotEqual = 4

    }
    export const enum YxFormatConditionType {
        xlAboveAverageCondition = 12,
        xlBlanksCondition = 10,
        xlCellValue = 1,
        xlColorScale = 3,
        xlDatabar = 4,
        xlErrorsCondition = 16,
        xlExpression = 2,
        xlIconSets = 6,
        xlNoBlanksCondition = 13,
        xlNoErrorsCondition = 17,
        xlTextString = 9,
        xlTimePeriod = 11,
        xlTop10 = 5,
        xlUniqueValues = 8

    }
    export const enum YxFormatConditionType {
        xlAboveAverageCondition = 12,
        xlBlanksCondition = 10,
        xlCellValue = 1,
        xlColorScale = 3,
        xlDatabar = 4,
        xlErrorsCondition = 16,
        xlExpression = 2,
        xlIconSets = 6,
        xlNoBlanksCondition = 13,
        xlNoErrorsCondition = 17,
        xlTextString = 9,
        xlTimePeriod = 11,
        xlTop10 = 5,
        xlUniqueValues = 8

    }
    export const enum YxFormControl {
        yxButtonControl = 0,
        yxCheckBox = 1,
        yxDropDown = 2,
        yxEditBox = 3,
        yxGroupBox = 4,
        yxLabel = 5,
        yxListBox = 6,
        yxOptionButton = 7,
        yxScrollBar = 8,
        yxSpinner = 9

    }
    export const enum YxFormControl {
        yxButtonControl = 0,
        yxCheckBox = 1,
        yxDropDown = 2,
        yxEditBox = 3,
        yxGroupBox = 4,
        yxLabel = 5,
        yxListBox = 6,
        yxOptionButton = 7,
        yxScrollBar = 8,
        yxSpinner = 9

    }
    export const enum YxFormulaLabel {
        yxColumnLabels = 2,
        yxMixedLabels = 3,
        yxNoLabels = -4142,
        yxRowLabels = 1

    }
    export const enum YxFormulaLabel {
        yxColumnLabels = 2,
        yxMixedLabels = 3,
        yxNoLabels = -4142,
        yxRowLabels = 1

    }
    export const enum YxHAlign {
        yxHAlignCenter = -4108,
        yxHAlignCenterAcrossSelection = 7,
        yxHAlignDistributed = -4117,
        yxHAlignFill = 5,
        yxHAlignGeneral = 1,
        yxHAlignJustify = -4130,
        yxHAlignLeft = -4131,
        yxHAlignRight = -4152

    }
    export const enum YxHAlign {
        yxHAlignCenter = -4108,
        yxHAlignCenterAcrossSelection = 7,
        yxHAlignDistributed = -4117,
        yxHAlignFill = 5,
        yxHAlignGeneral = 1,
        yxHAlignJustify = -4130,
        yxHAlignLeft = -4131,
        yxHAlignRight = -4152

    }
    export const enum YxHebrewModes {
        yxHebrewFullScript = 0,
        yxHebrewMixedAuthorizedScript = 3,
        yxHebrewMixedScript = 2,
        yxHebrewPartialScript = 1

    }
    export const enum YxHebrewModes {
        yxHebrewFullScript = 0,
        yxHebrewMixedAuthorizedScript = 3,
        yxHebrewMixedScript = 2,
        yxHebrewPartialScript = 1

    }
    export const enum YxHighlightChangesTime {
        yxAllChanges = 2,
        yxNotYetReviewed = 3,
        yxSinceMyLastSave = 1

    }
    export const enum YxHighlightChangesTime {
        yxAllChanges = 2,
        yxNotYetReviewed = 3,
        yxSinceMyLastSave = 1

    }
    export const enum YxIcon {
        xlIcon0Bars = 37,
        xlIcon0FilledBoxes = 52,
        xlIcon1Bar = 38,
        xlIcon1FilledBox = 51,
        xlIcon2Bars = 39,
        xlIcon2FilledBoxes = 50,
        xlIcon3Bars = 40,
        xlIcon3FilledBoxes = 49,
        xlIcon4Bars = 41,
        xlIcon4FilledBoxes = 48,
        xlIconBlackCircle = 32,
        xlIconBlackCircleWithBorder = 13,
        xlIconCircleWithOneWhiteQuarter = 33,
        xlIconCircleWithThreeWhiteQuarters = 35,
        xlIconCircleWithTwoWhiteQuarters = 34,
        xlIconGoldStar = 42,
        xlIconGrayCircle = 31,
        xlIconGrayDownArrow = 6,
        xlIconGrayDownInclineArrow = 28,
        xlIconGraySideArrow = 5,
        xlIconGrayUpArrow = 4,
        xlIconGrayUpInclineArrow = 27,
        xlIconGreenCheck = 22,
        xlIconGreenCheckSymbol = 19,
        xlIconGreenCircle = 10,
        xlIconGreenFlag = 7,
        xlIconGreenTrafficLight = 14,
        xlIconGreenUpArrow = 1,
        xlIconGreenUpTriangle = 45,
        xlIconHalfGoldStar = 43,
        xlIconNoCellIcon = -1,
        xlIconPinkCircle = 30,
        xlIconRedCircle = 29,
        xlIconRedCircleWithBorder = 12,
        xlIconRedCross = 24,
        xlIconRedCrossSymbol = 21,
        xlIconRedDiamond = 18,
        xlIconRedDownArrow = 3,
        xlIconRedDownTriangle = 47,
        xlIconRedFlag = 9,
        xlIconRedTrafficLight = 16,
        xlIconSilverStar = 44,
        xlIconWhiteCircleAllWhiteQuarters = 36,
        xlIconYellowCircle = 11,
        xlIconYellowDash = 46,
        xlIconYellowDownInclineArrow = 26,
        xlIconYellowExclamation = 23,
        xlIconYellowExclamationSymbol = 20,
        xlIconYellowFlag = 8,
        xlIconYellowSideArrow = 2,
        xlIconYellowTrafficLight = 15,
        xlIconYellowTriangle = 17,
        xlIconYellowUpInclineArrow = 25

    }
    export const enum YxIcon {
        xlIcon0Bars = 37,
        xlIcon0FilledBoxes = 52,
        xlIcon1Bar = 38,
        xlIcon1FilledBox = 51,
        xlIcon2Bars = 39,
        xlIcon2FilledBoxes = 50,
        xlIcon3Bars = 40,
        xlIcon3FilledBoxes = 49,
        xlIcon4Bars = 41,
        xlIcon4FilledBoxes = 48,
        xlIconBlackCircle = 32,
        xlIconBlackCircleWithBorder = 13,
        xlIconCircleWithOneWhiteQuarter = 33,
        xlIconCircleWithThreeWhiteQuarters = 35,
        xlIconCircleWithTwoWhiteQuarters = 34,
        xlIconGoldStar = 42,
        xlIconGrayCircle = 31,
        xlIconGrayDownArrow = 6,
        xlIconGrayDownInclineArrow = 28,
        xlIconGraySideArrow = 5,
        xlIconGrayUpArrow = 4,
        xlIconGrayUpInclineArrow = 27,
        xlIconGreenCheck = 22,
        xlIconGreenCheckSymbol = 19,
        xlIconGreenCircle = 10,
        xlIconGreenFlag = 7,
        xlIconGreenTrafficLight = 14,
        xlIconGreenUpArrow = 1,
        xlIconGreenUpTriangle = 45,
        xlIconHalfGoldStar = 43,
        xlIconNoCellIcon = -1,
        xlIconPinkCircle = 30,
        xlIconRedCircle = 29,
        xlIconRedCircleWithBorder = 12,
        xlIconRedCross = 24,
        xlIconRedCrossSymbol = 21,
        xlIconRedDiamond = 18,
        xlIconRedDownArrow = 3,
        xlIconRedDownTriangle = 47,
        xlIconRedFlag = 9,
        xlIconRedTrafficLight = 16,
        xlIconSilverStar = 44,
        xlIconWhiteCircleAllWhiteQuarters = 36,
        xlIconYellowCircle = 11,
        xlIconYellowDash = 46,
        xlIconYellowDownInclineArrow = 26,
        xlIconYellowExclamation = 23,
        xlIconYellowExclamationSymbol = 20,
        xlIconYellowFlag = 8,
        xlIconYellowSideArrow = 2,
        xlIconYellowTrafficLight = 15,
        xlIconYellowTriangle = 17,
        xlIconYellowUpInclineArrow = 25

    }
    export const enum YxIconSet {
        xl3Arrows = 1,
        xl3ArrowsGray = 2,
        xl3Flags = 3,
        xl3Signs = 6,
        xl3Stars = 18,
        xl3Symbols = 7,
        xl3Symbols2 = 8,
        xl3TrafficLights1 = 4,
        xl3TrafficLights2 = 5,
        xl3Triangles = 19,
        xl4Arrows = 9,
        xl4ArrowsGray = 10,
        xl4CRV = 12,
        xl4RedToBlack = 11,
        xl4TrafficLights = 13,
        xl5Arrows = 14,
        xl5ArrowsGray = 15,
        xl5Boxes = 20,
        xl5CRV = 16,
        xl5Quarters = 17,
        xlCustomSet = -1

    }
    export const enum YxIconSet {
        xl3Arrows = 1,
        xl3ArrowsGray = 2,
        xl3Flags = 3,
        xl3Signs = 6,
        xl3Stars = 18,
        xl3Symbols = 7,
        xl3Symbols2 = 8,
        xl3TrafficLights1 = 4,
        xl3TrafficLights2 = 5,
        xl3Triangles = 19,
        xl4Arrows = 9,
        xl4ArrowsGray = 10,
        xl4CRV = 12,
        xl4RedToBlack = 11,
        xl4TrafficLights = 13,
        xl5Arrows = 14,
        xl5ArrowsGray = 15,
        xl5Boxes = 20,
        xl5CRV = 16,
        xl5Quarters = 17,
        xlCustomSet = -1

    }
    export const enum YxIMEMode {
        yxIMEModeAlpha = 8,
        yxIMEModeAlphaFull = 7,
        yxIMEModeDisable = 3,
        yxIMEModeHangul = 10,
        yxIMEModeHangulFull = 9,
        yxIMEModeHiragana = 4,
        yxIMEModeKatakana = 5,
        yxIMEModeKatakanaHalf = 6,
        yxIMEModeNoControl = 0,
        yxIMEModeOff = 2,
        yxIMEModeOn = 1

    }
    export const enum YxIMEMode {
        yxIMEModeAlpha = 8,
        yxIMEModeAlphaFull = 7,
        yxIMEModeDisable = 3,
        yxIMEModeHangul = 10,
        yxIMEModeHangulFull = 9,
        yxIMEModeHiragana = 4,
        yxIMEModeKatakana = 5,
        yxIMEModeKatakanaHalf = 6,
        yxIMEModeNoControl = 0,
        yxIMEModeOff = 2,
        yxIMEModeOn = 1

    }
    export const enum YxInsertShiftDirection {
        yxShiftDown = -4121,
        yxShiftToRight = -4161

    }
    export const enum YxInsertShiftDirection {
        yxShiftDown = -4121,
        yxShiftToRight = -4161

    }
    export const enum YXLayoutFormType {
        yxOutline = 1,
        yxTabular = 0

    }
    export const enum YXLayoutFormType {
        yxOutline = 1,
        yxTabular = 0

    }
    export const enum YxLineStyle {
        yxContinuous = 1,
        yxDash = -4115,
        yxDashDot = 4,
        yxDashDotDot = 5,
        yxDot = -4118,
        yxDouble = -4119,
        yxLineStyleNone = -4142,
        yxSlantDashDot = 13

    }
    export const enum YxLineStyle {
        yxContinuous = 1,
        yxDash = -4115,
        yxDashDot = 4,
        yxDashDotDot = 5,
        yxDot = -4118,
        yxDouble = -4119,
        yxLineStyleNone = -4142,
        yxSlantDashDot = 13

    }
    export const enum YxLink {
        yxExcelLinks = 1,
        yxOLELinks = 2,
        yxPublishers = 5,
        yxSubscribers = 6

    }
    export const enum YxLink {
        yxExcelLinks = 1,
        yxOLELinks = 2,
        yxPublishers = 5,
        yxSubscribers = 6

    }
    export const enum YxLinkInfo {
        yxEditionDate = 2,
        yxLinkInfoStatus = 3,
        yxUpdateState = 1

    }
    export const enum YxLinkInfo {
        yxEditionDate = 2,
        yxLinkInfoStatus = 3,
        yxUpdateState = 1

    }
    export const enum YxLinkInfoType {
        yxLinkInfoOLELinks = 2,
        yxLinkInfoPublishers = 5,
        yxLinkInfoSubscribers = 6

    }
    export const enum YxLinkInfoType {
        yxLinkInfoOLELinks = 2,
        yxLinkInfoPublishers = 5,
        yxLinkInfoSubscribers = 6

    }
    export const enum YxLinkType {
        yxLinkTypeExcelLinks = 1,
        yxLinkTypeOLELinks = 2

    }
    export const enum YxLinkType {
        yxLinkTypeExcelLinks = 1,
        yxLinkTypeOLELinks = 2

    }
    export const enum YxLocationInTable {
        yxColumnHeader = -4110,
        yxColumnItem = 5,
        yxDataHeader = 3,
        yxDataItem = 7,
        yxPageHeader = 2,
        yxPageItem = 6,
        yxRowHeader = -4153,
        yxRowItem = 4,
        yxTableBody = 8

    }
    export const enum YxLocationInTable {
        yxColumnHeader = -4110,
        yxColumnItem = 5,
        yxDataHeader = 3,
        yxDataItem = 7,
        yxPageHeader = 2,
        yxPageItem = 6,
        yxRowHeader = -4153,
        yxRowItem = 4,
        yxTableBody = 8

    }
    export const enum YxLookAt {
        yxPart = 2,
        yxWhole = 1

    }
    export const enum YxLookAt {
        yxPart = 2,
        yxWhole = 1

    }
    export const enum YxMailSystem {
        yxMAPI = 1,
        yxNoMailSystem = 0,
        yxPowerTalk = 2

    }
    export const enum YxMailSystem {
        yxMAPI = 1,
        yxNoMailSystem = 0,
        yxPowerTalk = 2

    }
    export const enum YxMarkerStyle {
        yxMarkerStyleAutomatic = -4105,
        yxMarkerStyleCircle = 8,
        yxMarkerStyleDash = -4115,
        yxMarkerStyleDiamond = 2,
        yxMarkerStyleDot = -4118,
        yxMarkerStyleNone = -4142,
        yxMarkerStylePicture = -4147,
        yxMarkerStylePlus = 9,
        yxMarkerStyleSquare = 1,
        yxMarkerStyleStar = 5,
        yxMarkerStyleTriangle = 3,
        yxMarkerStyleX = -4168

    }
    export const enum YxMarkerStyle {
        yxMarkerStyleAutomatic = -4105,
        yxMarkerStyleCircle = 8,
        yxMarkerStyleDash = -4115,
        yxMarkerStyleDiamond = 2,
        yxMarkerStyleDot = -4118,
        yxMarkerStyleNone = -4142,
        yxMarkerStylePicture = -4147,
        yxMarkerStylePlus = 9,
        yxMarkerStyleSquare = 1,
        yxMarkerStyleStar = 5,
        yxMarkerStyleTriangle = 3,
        yxMarkerStyleX = -4168

    }
    export const enum YxMenuID {
        yxFile = 30002,
        yxEdit = 30003,
        yxView = 30004,
        yxInsert = 30005,
        yxFormat = 30006,
        yxTools = 30007,
        yxData = 30011,
        yxWindow = 30009,
        yxHelp = 30010,
        yxFileNew = 18,
        yxFileOpen = 23,
        yxFileClose = 106,
        yxFileSave = 3,
        yxFileSaveAs = 748,
        yxFileSaveAsWebPage = 3823,
        yxFileSaveWorkSpace = 846,
        yxFileWebPagePreview = 3655,
        yxFilePageSetup = 247,
        yxFilePrintArea = 30255,
        yxFilePrintPreivew = 109,
        yxFilePrint = 4,
        yxFileSendTo = 30095,
        yxFileProPerties = 750,
        yxFileRecentFileName = 831,
        yxFileExit = 752,
        yxEditUndo = 128,
        yxEditRepeat = 37,
        yxEditCut = 21,
        yxEditCopy = 19,
        yxEditPaste = 22,
        yxEditPasteSpecial = 755,
        yxEditPasteHyperlink = 2787,
        yxEditFill = 30020,
        yxEditFillDown = 372,
        yxEditFillRight = 371,
        yxEditFillUp = 867,
        yxEditFillLeft = 868,
        yxEditFillAcrossWorksheets = 869,
        yxEditFillSeries = 870,
        yxEditFillJustify = 871,
        yxEditClear = 30021,
        yxEditClearAll = 1964,
        yxEditClearFormats = 872,
        yxEditClearContents = 873,
        yxEditClearComments = 874,
        yxEditDelete = 478,
        yxEditDeleteSheet = 847,
        yxEditMoveOrCopySheet = 848,
        yxEditFind = 1849,
        yxEditReplace = 313,
        yxEditGoto = 757,
        yxEditLinks = 759,
        yxViewNormal = 723,
        yxViewPageBreakPreivew = 724,
        yxViewToolbars = 30045,
        yxViewFormulaBar = 849,
        yxViewStatusBar = 850,
        yxViewHeaderAndFooter = 762,
        yxViewComments = 1594,
        yxViewCustomViews = 950,
        yxViewFullScreen = 178,
        yxViewZoom = 925,
        yxInsertCells = 295,
        yxInsertRows = 296,
        yxInsertColumns = 297,
        yxInsertWorksheet = 852,
        yxInsertChart = 1957,
        yxInsertSymbol = 308,
        yxInsertPageBreak = 509,
        yxInsertResetAllPageBreaks = 1585,
        yxInsertFunction = 385,
        yxInsertName = 30023,
        yxInsertNameDefine = 878,
        yxInsertNamePaste = 879,
        yxInsertNameCreate = 880,
        yxInsertNameApply = 881,
        yxInsertNameLabel = 932,
        yxInsertComment = 1589,
        yxInsertPicture = 30180,
        yxInsertPictureFromFile = 2619,
        yxInsertHyperlink = 1576,
        yxFormatCells = 855,
        yxFormatRow = 30024,
        yxFormatRowHeight = 541,
        yxFormatRowHeightAutoFit = 882,
        yxFormatRowHide = 883,
        yxFormatRowUnhide = 884,
        yxFormatColumn = 30025,
        yxFormatColumnWidth = 542,
        yxFormatColumnWidthAutoFit = 885,
        yxFormatColumnsHide = 886,
        yxFormatColumnsUnhide = 887,
        yxFormatColumnWidthDefault = 888,
        yxFormatSheet = 30026,
        yxFormatSheetRename = 889,
        yxFormatSheetHide = 890,
        yxFormatSheetUnhide = 891,
        yxFormatAutoFormat = 786,
        yxFormatConditionalFormatting = 3058,
        yxFormatStyle = 254,
        yxFormatPhoneictGuide = 30136,
        yxToolsSpelling = 2,
        yxToolsAutoCorrect = 793,
        yxToolsShareWorkbook = 2040,
        yxToolsTrackChanges = 30138,
        yxToolsHighlightChanges = 2042,
        yxToolsAcceptOrRejectChange = 305,
        yxToolsMergeWorkbooks = 2044,
        yxToolsProtection = 30029,
        yxToolsProtectSheet = 893,
        yxToolsAllowUsersToEditRanges = 6997,
        yxToolsProtectWorkbook = 894,
        yxToolsProtectAndShareWorkbook = 3059,
        yxToolsOnlineCollaboration = 30468,
        yxToolsGoalSeek = 856,
        yxToolsScenarios = 857,
        yxToolsAuditing = 30028,
        yxToolsAuditPrecedents = 486,
        yxToolsAuditDependents = 451,
        yxToolsAuditError = 463,
        yxToolsAuditRemoveAllArrows = 453,
        yxToolsAuditFormulaEvaluate = 5687,
        yxToolsAuditWatchWindow = 5686,
        yxToolsAuditShowFormulas = 6059,
        yxToolsAuditToolbar = 892,
        yxToolsMacro = 30017,
        yxToolsAddins = 943,
        yxToolsCustomize = 797,
        yxToolsOptions = 522,
        yxDataSort = 928,
        yxDataFilter = 30031,
        yxDataAutoFilter = 899,
        yxDataSortShow = 900,
        yxDataAdvancedFilter = 901,
        yxDataForm = 860,
        yxDataSubtotals = 861,
        yxDataValidation = 2034,
        yxDataTable = 862,
        yxDataTextToColumns = 806,
        yxDataConsolidate = 863,
        yxDataGroupAndOutline = 30032,
        yxDataOutlineHideDetail = 464,
        yxDataOutlineShowDetail = 462,
        yxDataOutlineGroup = 3159,
        yxDataOutlineUngroup = 3160,
        yxDataOutlineAuto = 904,
        yxDataOutlineClear = 905,
        yxDataOutlineSettings = 906,
        yxDataPivoTable = 2915,
        yxDataGetExternalData = 30101,
        yxDataImportClassic = 6262,
        yxDataFromWeb = 3829,
        yxDataFromDatabase = 2054,
        yxDataEditQuery = 1950,
        yxDataRangeProperties = 1951,
        yxDataQueryParameters = 537,
        yxDataTableList = 31276,
        yxDataTableInsertExcel = 7193,
        yxDataTableResize = 7765,
        yxDataTableStyleTotalsRow = 7372,
        yxDataTableConvertToRange = 7375,
        yxDataTableToSharePointList = 7684,
        yxDataTableOpenInBrowser = 7763,
        yxDataTableUnlinkExternalData = 7683,
        yxDataTableListSynchronize = 7761,
        yxDataTableChangesDiscardAndRefresh = 7762,
        yxDataTableHideDeactiveListBorder = 9747,
        yxDataXML = 31268,
        yxDataXMLImport = 7433,
        yxDataXMLExmport = 7432,
        yxDataXMLDataRefresh = 7812,
        yxDataXMLSource = 7693,
        yxDataXmlMapProperties = 7813,
        yxDataXMLEditQuery = 7897,
        yxDataXmlExpansionPacksExcel = 7784,
        yxDataRefreshData = 459,
        yxWindowNewWindow = 303,
        yxWindowArrange = 298,
        yxWindowHide = 865,
        yxWindowUnhide = 866,
        yxWindowSplit = 302,
        yxWindowFreezePanes = 443,
        yxStandardNew = 2520,
        yxStandardOpen = 23,
        yxStandardSave = 3,
        yxStandardMailRecipient = 3738,
        yxStandardPrint = 2521,
        yxStandardPrintPreview = 109,
        yxStandardSpelling = 2,
        yxStandardCut = 21,
        yxStandardCopy = 19,
        yxStandardPaste = 6002,
        yxStandardFormatPainter = 108,
        yxStandardUndo = 128,
        yxStandardRedo = 129,
        yxStandardHyperlink = 1576,
        yxStandardAutoSum = 226,
        yxStandardPasteFunction = 385,
        yxStandardSortAscending = 210,
        yxStandardSortDescending = 211,
        yxStandardChartWizard = 436,
        yxStandardDrawing = 204,
        yxStandardZoom = 1733,
        yxStandardAutoFilter = 458,
        yxStandardHelp = 984,
        yxFormattingFont = 1728,
        yxFormattingFontSize = 1731,
        yxFormattingBold = 113,
        yxFormattingItalic = 114,
        yxFormattingUnderline = 115,
        yxFormattingAlignLeft = 120,
        yxFormattingCenter = 122,
        yxFormattingAlignRight = 121,
        yxFormattingMergeAndCenter = 402,
        yxFormattingCurrencyStyle = 395,
        yxFormattingPercentStyle = 396,
        yxFormattingCommaStyle = 397,
        yxFormattingIncreaseDecimal = 398,
        yxFormattingDecreaseDecimal = 399,
        yxFormattingDecreaseIndent = 3162,
        yxFormattingIncreaseIndent = 3161,
        yxFormattingBorders = 203,
        yxFormattingFillColor = 1691,
        yxFormattingFontColor = 401,
        yxFormattingFontSizeInscrease = 403,
        yxFormattingFontSizeDescrease = 404,
        yxReviewNewComment = 1589,
        yxReviewPreviousComment = 1590,
        yxReviewNextComment = 1591,
        yxReviewShowOrHideComment = 1593,
        yxReviewShowAllComments = 1594,
        yxReviewDeleteComment = 1592,
        yxReviewInkingStart = 9071,
        yxReviewShowInk = 7500,
        yxReviewDeleteAllInk = 7674,
        yxReviewEndReview = 6141

    }
    export const enum YxMenuID {
        yxFile = 30002,
        yxEdit = 30003,
        yxView = 30004,
        yxInsert = 30005,
        yxFormat = 30006,
        yxTools = 30007,
        yxData = 30011,
        yxWindow = 30009,
        yxHelp = 30010,
        yxFileNew = 18,
        yxFileOpen = 23,
        yxFileClose = 106,
        yxFileSave = 3,
        yxFileSaveAs = 748,
        yxFileSaveAsWebPage = 3823,
        yxFileSaveWorkSpace = 846,
        yxFileWebPagePreview = 3655,
        yxFilePageSetup = 247,
        yxFilePrintArea = 30255,
        yxFilePrintPreivew = 109,
        yxFilePrint = 4,
        yxFileSendTo = 30095,
        yxFileProPerties = 750,
        yxFileRecentFileName = 831,
        yxFileExit = 752,
        yxEditUndo = 128,
        yxEditRepeat = 37,
        yxEditCut = 21,
        yxEditCopy = 19,
        yxEditPaste = 22,
        yxEditPasteSpecial = 755,
        yxEditPasteHyperlink = 2787,
        yxEditFill = 30020,
        yxEditFillDown = 372,
        yxEditFillRight = 371,
        yxEditFillUp = 867,
        yxEditFillLeft = 868,
        yxEditFillAcrossWorksheets = 869,
        yxEditFillSeries = 870,
        yxEditFillJustify = 871,
        yxEditClear = 30021,
        yxEditClearAll = 1964,
        yxEditClearFormats = 872,
        yxEditClearContents = 873,
        yxEditClearComments = 874,
        yxEditDelete = 478,
        yxEditDeleteSheet = 847,
        yxEditMoveOrCopySheet = 848,
        yxEditFind = 1849,
        yxEditReplace = 313,
        yxEditGoto = 757,
        yxEditLinks = 759,
        yxViewNormal = 723,
        yxViewPageBreakPreivew = 724,
        yxViewToolbars = 30045,
        yxViewFormulaBar = 849,
        yxViewStatusBar = 850,
        yxViewHeaderAndFooter = 762,
        yxViewComments = 1594,
        yxViewCustomViews = 950,
        yxViewFullScreen = 178,
        yxViewZoom = 925,
        yxInsertCells = 295,
        yxInsertRows = 296,
        yxInsertColumns = 297,
        yxInsertWorksheet = 852,
        yxInsertChart = 1957,
        yxInsertSymbol = 308,
        yxInsertPageBreak = 509,
        yxInsertResetAllPageBreaks = 1585,
        yxInsertFunction = 385,
        yxInsertName = 30023,
        yxInsertNameDefine = 878,
        yxInsertNamePaste = 879,
        yxInsertNameCreate = 880,
        yxInsertNameApply = 881,
        yxInsertNameLabel = 932,
        yxInsertComment = 1589,
        yxInsertPicture = 30180,
        yxInsertPictureFromFile = 2619,
        yxInsertHyperlink = 1576,
        yxFormatCells = 855,
        yxFormatRow = 30024,
        yxFormatRowHeight = 541,
        yxFormatRowHeightAutoFit = 882,
        yxFormatRowHide = 883,
        yxFormatRowUnhide = 884,
        yxFormatColumn = 30025,
        yxFormatColumnWidth = 542,
        yxFormatColumnWidthAutoFit = 885,
        yxFormatColumnsHide = 886,
        yxFormatColumnsUnhide = 887,
        yxFormatColumnWidthDefault = 888,
        yxFormatSheet = 30026,
        yxFormatSheetRename = 889,
        yxFormatSheetHide = 890,
        yxFormatSheetUnhide = 891,
        yxFormatAutoFormat = 786,
        yxFormatConditionalFormatting = 3058,
        yxFormatStyle = 254,
        yxFormatPhoneictGuide = 30136,
        yxToolsSpelling = 2,
        yxToolsAutoCorrect = 793,
        yxToolsShareWorkbook = 2040,
        yxToolsTrackChanges = 30138,
        yxToolsHighlightChanges = 2042,
        yxToolsAcceptOrRejectChange = 305,
        yxToolsMergeWorkbooks = 2044,
        yxToolsProtection = 30029,
        yxToolsProtectSheet = 893,
        yxToolsAllowUsersToEditRanges = 6997,
        yxToolsProtectWorkbook = 894,
        yxToolsProtectAndShareWorkbook = 3059,
        yxToolsOnlineCollaboration = 30468,
        yxToolsGoalSeek = 856,
        yxToolsScenarios = 857,
        yxToolsAuditing = 30028,
        yxToolsAuditPrecedents = 486,
        yxToolsAuditDependents = 451,
        yxToolsAuditError = 463,
        yxToolsAuditRemoveAllArrows = 453,
        yxToolsAuditFormulaEvaluate = 5687,
        yxToolsAuditWatchWindow = 5686,
        yxToolsAuditShowFormulas = 6059,
        yxToolsAuditToolbar = 892,
        yxToolsMacro = 30017,
        yxToolsAddins = 943,
        yxToolsCustomize = 797,
        yxToolsOptions = 522,
        yxDataSort = 928,
        yxDataFilter = 30031,
        yxDataAutoFilter = 899,
        yxDataSortShow = 900,
        yxDataAdvancedFilter = 901,
        yxDataForm = 860,
        yxDataSubtotals = 861,
        yxDataValidation = 2034,
        yxDataTable = 862,
        yxDataTextToColumns = 806,
        yxDataConsolidate = 863,
        yxDataGroupAndOutline = 30032,
        yxDataOutlineHideDetail = 464,
        yxDataOutlineShowDetail = 462,
        yxDataOutlineGroup = 3159,
        yxDataOutlineUngroup = 3160,
        yxDataOutlineAuto = 904,
        yxDataOutlineClear = 905,
        yxDataOutlineSettings = 906,
        yxDataPivoTable = 2915,
        yxDataGetExternalData = 30101,
        yxDataImportClassic = 6262,
        yxDataFromWeb = 3829,
        yxDataFromDatabase = 2054,
        yxDataEditQuery = 1950,
        yxDataRangeProperties = 1951,
        yxDataQueryParameters = 537,
        yxDataTableList = 31276,
        yxDataTableInsertExcel = 7193,
        yxDataTableResize = 7765,
        yxDataTableStyleTotalsRow = 7372,
        yxDataTableConvertToRange = 7375,
        yxDataTableToSharePointList = 7684,
        yxDataTableOpenInBrowser = 7763,
        yxDataTableUnlinkExternalData = 7683,
        yxDataTableListSynchronize = 7761,
        yxDataTableChangesDiscardAndRefresh = 7762,
        yxDataTableHideDeactiveListBorder = 9747,
        yxDataXML = 31268,
        yxDataXMLImport = 7433,
        yxDataXMLExmport = 7432,
        yxDataXMLDataRefresh = 7812,
        yxDataXMLSource = 7693,
        yxDataXmlMapProperties = 7813,
        yxDataXMLEditQuery = 7897,
        yxDataXmlExpansionPacksExcel = 7784,
        yxDataRefreshData = 459,
        yxWindowNewWindow = 303,
        yxWindowArrange = 298,
        yxWindowHide = 865,
        yxWindowUnhide = 866,
        yxWindowSplit = 302,
        yxWindowFreezePanes = 443,
        yxStandardNew = 2520,
        yxStandardOpen = 23,
        yxStandardSave = 3,
        yxStandardMailRecipient = 3738,
        yxStandardPrint = 2521,
        yxStandardPrintPreview = 109,
        yxStandardSpelling = 2,
        yxStandardCut = 21,
        yxStandardCopy = 19,
        yxStandardPaste = 6002,
        yxStandardFormatPainter = 108,
        yxStandardUndo = 128,
        yxStandardRedo = 129,
        yxStandardHyperlink = 1576,
        yxStandardAutoSum = 226,
        yxStandardPasteFunction = 385,
        yxStandardSortAscending = 210,
        yxStandardSortDescending = 211,
        yxStandardChartWizard = 436,
        yxStandardDrawing = 204,
        yxStandardZoom = 1733,
        yxStandardAutoFilter = 458,
        yxStandardHelp = 984,
        yxFormattingFont = 1728,
        yxFormattingFontSize = 1731,
        yxFormattingBold = 113,
        yxFormattingItalic = 114,
        yxFormattingUnderline = 115,
        yxFormattingAlignLeft = 120,
        yxFormattingCenter = 122,
        yxFormattingAlignRight = 121,
        yxFormattingMergeAndCenter = 402,
        yxFormattingCurrencyStyle = 395,
        yxFormattingPercentStyle = 396,
        yxFormattingCommaStyle = 397,
        yxFormattingIncreaseDecimal = 398,
        yxFormattingDecreaseDecimal = 399,
        yxFormattingDecreaseIndent = 3162,
        yxFormattingIncreaseIndent = 3161,
        yxFormattingBorders = 203,
        yxFormattingFillColor = 1691,
        yxFormattingFontColor = 401,
        yxFormattingFontSizeInscrease = 403,
        yxFormattingFontSizeDescrease = 404,
        yxReviewNewComment = 1589,
        yxReviewPreviousComment = 1590,
        yxReviewNextComment = 1591,
        yxReviewShowOrHideComment = 1593,
        yxReviewShowAllComments = 1594,
        yxReviewDeleteComment = 1592,
        yxReviewInkingStart = 9071,
        yxReviewShowInk = 7500,
        yxReviewDeleteAllInk = 7674,
        yxReviewEndReview = 6141

    }
    export const enum YxMousePointer {
        yxDefault = -4143,
        yxIBeam = 3,
        yxNorthwestArrow = 1,
        yxWait = 2

    }
    export const enum YxMousePointer {
        yxDefault = -4143,
        yxIBeam = 3,
        yxNorthwestArrow = 1,
        yxWait = 2

    }
    export const enum YxMSApplication {
        yxMicrosoftAccess = 4,
        yxMicrosoftFoxPro = 5,
        yxMicrosoftMail = 3,
        yxMicrosoftPowerPoint = 2,
        yxMicrosoftProject = 6,
        yxMicrosoftSchedulePlus = 7,
        yxMicrosoftWord = 1

    }
    export const enum YxMSApplication {
        yxMicrosoftAccess = 4,
        yxMicrosoftFoxPro = 5,
        yxMicrosoftMail = 3,
        yxMicrosoftPowerPoint = 2,
        yxMicrosoftProject = 6,
        yxMicrosoftSchedulePlus = 7,
        yxMicrosoftWord = 1

    }
    export const enum YxOrder {
        yxDownThenOver = 1,
        yxOverThenDown = 2

    }
    export const enum YxOrder {
        yxDownThenOver = 1,
        yxOverThenDown = 2

    }
    export const enum YxOrientation {
        yxDownward = -4170,
        yxHorizontal = -4128,
        yxUpward = -4171,
        yxVertical = -4166

    }
    export const enum YxOrientation {
        yxDownward = -4170,
        yxHorizontal = -4128,
        yxUpward = -4171,
        yxVertical = -4166

    }
    export const enum YxPageBreak {
        yxPageBreakAutomatic = -4105,
        yxPageBreakManual = -4135,
        yxPageBreakNone = -4142

    }
    export const enum YxPageBreak {
        yxPageBreakAutomatic = -4105,
        yxPageBreakManual = -4135,
        yxPageBreakNone = -4142

    }
    export const enum YxPageBreakExtent {
        yxPageBreakFull = 1,
        yxPageBreakPartial = 2

    }
    export const enum YxPageBreakExtent {
        yxPageBreakFull = 1,
        yxPageBreakPartial = 2

    }
    export const enum YxPageOrientation {
        yxLandscape = 2,
        yxPortrait = 1

    }
    export const enum YxPageOrientation {
        yxLandscape = 2,
        yxPortrait = 1

    }
    export const enum YxPaperSize {
        yxPaper10x14 = 16,
        yxPaper11x17 = 17,
        yxPaperA3 = 8,
        yxPaperA4 = 9,
        yxPaperA4Small = 10,
        yxPaperA5 = 11,
        yxPaperB4 = 12,
        yxPaperB5 = 13,
        yxPaperCSheet = 24,
        yxPaperDSheet = 25,
        yxPaperEnvelope10 = 20,
        yxPaperEnvelope11 = 21,
        yxPaperEnvelope12 = 22,
        yxPaperEnvelope14 = 23,
        yxPaperEnvelope9 = 19,
        yxPaperEnvelopeB4 = 33,
        yxPaperEnvelopeB5 = 34,
        yxPaperEnvelopeB6 = 35,
        yxPaperEnvelopeC3 = 29,
        yxPaperEnvelopeC4 = 30,
        yxPaperEnvelopeC5 = 28,
        yxPaperEnvelopeC6 = 31,
        yxPaperEnvelopeC65 = 32,
        yxPaperEnvelopeDL = 27,
        yxPaperEnvelopeItaly = 36,
        yxPaperEnvelopeMonarch = 37,
        yxPaperEnvelopePersonal = 38,
        yxPaperESheet = 26,
        yxPaperExecutive = 7,
        yxPaperFanfoldLegalGerman = 41,
        yxPaperFanfoldStdGerman = 40,
        yxPaperFanfoldUS = 39,
        yxPaperFolio = 14,
        yxPaperLedger = 4,
        yxPaperLegal = 5,
        yxPaperLetter = 1,
        yxPaperLetterSmall = 2,
        yxPaperNote = 18,
        yxPaperQuarto = 15,
        yxPaperStatement = 6,
        yxPaperTabloid = 3,
        yxPaperUser = 256

    }
    export const enum YxPaperSize {
        yxPaper10x14 = 16,
        yxPaper11x17 = 17,
        yxPaperA3 = 8,
        yxPaperA4 = 9,
        yxPaperA4Small = 10,
        yxPaperA5 = 11,
        yxPaperB4 = 12,
        yxPaperB5 = 13,
        yxPaperCSheet = 24,
        yxPaperDSheet = 25,
        yxPaperEnvelope10 = 20,
        yxPaperEnvelope11 = 21,
        yxPaperEnvelope12 = 22,
        yxPaperEnvelope14 = 23,
        yxPaperEnvelope9 = 19,
        yxPaperEnvelopeB4 = 33,
        yxPaperEnvelopeB5 = 34,
        yxPaperEnvelopeB6 = 35,
        yxPaperEnvelopeC3 = 29,
        yxPaperEnvelopeC4 = 30,
        yxPaperEnvelopeC5 = 28,
        yxPaperEnvelopeC6 = 31,
        yxPaperEnvelopeC65 = 32,
        yxPaperEnvelopeDL = 27,
        yxPaperEnvelopeItaly = 36,
        yxPaperEnvelopeMonarch = 37,
        yxPaperEnvelopePersonal = 38,
        yxPaperESheet = 26,
        yxPaperExecutive = 7,
        yxPaperFanfoldLegalGerman = 41,
        yxPaperFanfoldStdGerman = 40,
        yxPaperFanfoldUS = 39,
        yxPaperFolio = 14,
        yxPaperLedger = 4,
        yxPaperLegal = 5,
        yxPaperLetter = 1,
        yxPaperLetterSmall = 2,
        yxPaperNote = 18,
        yxPaperQuarto = 15,
        yxPaperStatement = 6,
        yxPaperTabloid = 3,
        yxPaperUser = 256

    }
    export const enum YxPasteSpecialOperation {
        yxPasteSpecialOperationAdd = 2,
        yxPasteSpecialOperationDivide = 5,
        yxPasteSpecialOperationMultiply = 4,
        yxPasteSpecialOperationNone = -4142,
        yxPasteSpecialOperationSubtract = 3

    }
    export const enum YxPasteSpecialOperation {
        yxPasteSpecialOperationAdd = 2,
        yxPasteSpecialOperationDivide = 5,
        yxPasteSpecialOperationMultiply = 4,
        yxPasteSpecialOperationNone = -4142,
        yxPasteSpecialOperationSubtract = 3

    }
    export const enum YxPasteType {
        yxPasteAll = -4104,
        yxPasteAllExceptBorders = 7,
        yxPasteColumnWidths = 8,
        yxPasteComments = -4144,
        yxPasteFormats = -4122,
        yxPasteFormulas = -4123,
        yxPasteFormulasAndNumberFormats = 11,
        yxPasteValidation = 6,
        yxPasteValues = -4163,
        yxPasteValuesAndNumberFormats = 12

    }
    export const enum YxPasteType {
        yxPasteAll = -4104,
        yxPasteAllExceptBorders = 7,
        yxPasteColumnWidths = 8,
        yxPasteComments = -4144,
        yxPasteFormats = -4122,
        yxPasteFormulas = -4123,
        yxPasteFormulasAndNumberFormats = 11,
        yxPasteValidation = 6,
        yxPasteValues = -4163,
        yxPasteValuesAndNumberFormats = 12

    }
    export const enum YxPictureAppearance {
        yxPrinter = 2,
        yxScreen = 1

    }
    export const enum YxPictureAppearance {
        yxPrinter = 2,
        yxScreen = 1

    }
    export const enum YxPivotFieldOrientation {
        yxColumnField = 2,
        yxDataField = 4,
        yxHidden = 0,
        yxPageField = 3,
        yxRowField = 1

    }
    export const enum YxPivotFieldOrientation {
        yxColumnField = 2,
        yxDataField = 4,
        yxHidden = 0,
        yxPageField = 3,
        yxRowField = 1

    }
    export const enum YxPivotTableSourceType {
        yxConsolidation = 3,
        yxDatabase = 1,
        yxExternal = 2,
        yxPivotTable = -4148,
        yxScenario = 4

    }
    export const enum YxPivotTableSourceType {
        yxConsolidation = 3,
        yxDatabase = 1,
        yxExternal = 2,
        yxPivotTable = -4148,
        yxScenario = 4

    }
    export const enum YxPlacement {
        yxFreeFloating = 3,
        yxMove = 2,
        yxMoveAndSize = 1

    }
    export const enum YxPlacement {
        yxFreeFloating = 3,
        yxMove = 2,
        yxMoveAndSize = 1

    }
    export const enum YxPlatform {
        yxMacintosh = 1,
        yxMSDOS = 3,
        yxWindows = 2

    }
    export const enum YxPlatform {
        yxMacintosh = 1,
        yxMSDOS = 3,
        yxWindows = 2

    }
    export const enum YxPrintErrors {
        yxPrintErrorsBlank = 1,
        yxPrintErrorsDash = 2,
        yxPrintErrorsDisplayed = 0,
        yxPrintErrorsNA = 3

    }
    export const enum YxPrintErrors {
        yxPrintErrorsBlank = 1,
        yxPrintErrorsDash = 2,
        yxPrintErrorsDisplayed = 0,
        yxPrintErrorsNA = 3

    }
    export const enum YxPrintLocation {
        yxPrintInPlace = 16,
        yxPrintNoComments = -4142,
        yxPrintSheetEnd = 1

    }
    export const enum YxPrintLocation {
        yxPrintInPlace = 16,
        yxPrintNoComments = -4142,
        yxPrintSheetEnd = 1

    }
    export const enum YxPriority {
        yxPriorityHigh = -4127,
        yxPriorityLow = -4134,
        yxPriorityNormal = -4143

    }
    export const enum YxPriority {
        yxPriorityHigh = -4127,
        yxPriorityLow = -4134,
        yxPriorityNormal = -4143

    }
    export const enum YxRangeAutoFormat {
        yxRangeAutoFormat3DEffects1 = 13,
        yxRangeAutoFormat3DEffects2 = 14,
        yxRangeAutoFormatAccounting1 = 4,
        yxRangeAutoFormatAccounting2 = 5,
        yxRangeAutoFormatAccounting3 = 6,
        yxRangeAutoFormatAccounting4 = 17,
        yxRangeAutoFormatClassic1 = 1,
        yxRangeAutoFormatClassic2 = 2,
        yxRangeAutoFormatClassic3 = 3,
        yxRangeAutoFormatClassicPivotTable = 31,
        yxRangeAutoFormatColor1 = 7,
        yxRangeAutoFormatColor2 = 8,
        yxRangeAutoFormatColor3 = 9,
        yxRangeAutoFormatList1 = 10,
        yxRangeAutoFormatList2 = 11,
        yxRangeAutoFormatList3 = 12,
        yxRangeAutoFormatLocalFormat1 = 15,
        yxRangeAutoFormatLocalFormat2 = 16,
        yxRangeAutoFormatLocalFormat3 = 19,
        yxRangeAutoFormatLocalFormat4 = 20,
        yxRangeAutoFormatNone = -4142,
        yxRangeAutoFormatPTNone = 42,
        yxRangeAutoFormatReport1 = 21,
        yxRangeAutoFormatReport2 = 22,
        yxRangeAutoFormatReport3 = 23,
        yxRangeAutoFormatReport4 = 24,
        yxRangeAutoFormatReport5 = 25,
        yxRangeAutoFormatReport6 = 26,
        yxRangeAutoFormatReport7 = 27,
        yxRangeAutoFormatReport8 = 28,
        yxRangeAutoFormatReport9 = 29,
        yxRangeAutoFormatReport10 = 30,
        yxRangeAutoFormatSimple = -4154,
        yxRangeAutoFormatTable1 = 32,
        yxRangeAutoFormatTable2 = 33,
        yxRangeAutoFormatTable3 = 34,
        yxRangeAutoFormatTable4 = 35,
        yxRangeAutoFormatTable5 = 36,
        yxRangeAutoFormatTable6 = 37,
        yxRangeAutoFormatTable7 = 38,
        yxRangeAutoFormatTable8 = 39,
        yxRangeAutoFormatTable9 = 40,
        yxRangeAutoFormatTable10 = 41

    }
    export const enum YxRangeAutoFormat {
        yxRangeAutoFormat3DEffects1 = 13,
        yxRangeAutoFormat3DEffects2 = 14,
        yxRangeAutoFormatAccounting1 = 4,
        yxRangeAutoFormatAccounting2 = 5,
        yxRangeAutoFormatAccounting3 = 6,
        yxRangeAutoFormatAccounting4 = 17,
        yxRangeAutoFormatClassic1 = 1,
        yxRangeAutoFormatClassic2 = 2,
        yxRangeAutoFormatClassic3 = 3,
        yxRangeAutoFormatClassicPivotTable = 31,
        yxRangeAutoFormatColor1 = 7,
        yxRangeAutoFormatColor2 = 8,
        yxRangeAutoFormatColor3 = 9,
        yxRangeAutoFormatList1 = 10,
        yxRangeAutoFormatList2 = 11,
        yxRangeAutoFormatList3 = 12,
        yxRangeAutoFormatLocalFormat1 = 15,
        yxRangeAutoFormatLocalFormat2 = 16,
        yxRangeAutoFormatLocalFormat3 = 19,
        yxRangeAutoFormatLocalFormat4 = 20,
        yxRangeAutoFormatNone = -4142,
        yxRangeAutoFormatPTNone = 42,
        yxRangeAutoFormatReport1 = 21,
        yxRangeAutoFormatReport2 = 22,
        yxRangeAutoFormatReport3 = 23,
        yxRangeAutoFormatReport4 = 24,
        yxRangeAutoFormatReport5 = 25,
        yxRangeAutoFormatReport6 = 26,
        yxRangeAutoFormatReport7 = 27,
        yxRangeAutoFormatReport8 = 28,
        yxRangeAutoFormatReport9 = 29,
        yxRangeAutoFormatReport10 = 30,
        yxRangeAutoFormatSimple = -4154,
        yxRangeAutoFormatTable1 = 32,
        yxRangeAutoFormatTable2 = 33,
        yxRangeAutoFormatTable3 = 34,
        yxRangeAutoFormatTable4 = 35,
        yxRangeAutoFormatTable5 = 36,
        yxRangeAutoFormatTable6 = 37,
        yxRangeAutoFormatTable7 = 38,
        yxRangeAutoFormatTable8 = 39,
        yxRangeAutoFormatTable9 = 40,
        yxRangeAutoFormatTable10 = 41

    }
    export const enum YxRangeValueDataType {
        yxRangeValueDefault = 10,
        yxRangeValueMSPersistXML = 12,
        yxRangeValueXMLSpreadsheet = 11

    }
    export const enum YxRangeValueDataType {
        yxRangeValueDefault = 10,
        yxRangeValueMSPersistXML = 12,
        yxRangeValueXMLSpreadsheet = 11

    }
    export const enum YxReferenceStyle {
        yxA1 = 1,
        yxR1C1 = -4150

    }
    export const enum YxReferenceStyle {
        yxA1 = 1,
        yxR1C1 = -4150

    }
    export const enum YxRoutingSlipDelivery {
        yxAllAtOnce = 2,
        yxOneAfterAnother = 1

    }
    export const enum YxRoutingSlipDelivery {
        yxAllAtOnce = 2,
        yxOneAfterAnother = 1

    }
    export const enum YxRoutingSlipStatus {
        yxNotYetRouted = 0,
        yxRoutingComplete = 2,
        yxRoutingInProgress = 1

    }
    export const enum YxRoutingSlipStatus {
        yxNotYetRouted = 0,
        yxRoutingComplete = 2,
        yxRoutingInProgress = 1

    }
    export const enum YxRowCol {
        yxColumns = 2,
        yxRows = 1

    }
    export const enum YxRowCol {
        yxColumns = 2,
        yxRows = 1

    }
    export const enum YxRunAutoMacro {
        yxAutoActivate = 3,
        yxAutoClose = 2,
        yxAutoDeactivate = 4,
        yxAutoOpen = 1

    }
    export const enum YxRunAutoMacro {
        yxAutoActivate = 3,
        yxAutoClose = 2,
        yxAutoDeactivate = 4,
        yxAutoOpen = 1

    }
    export const enum YxSaveAction {
        yxDoNotSaveChanges = 2,
        yxSaveChanges = 1

    }
    export const enum YxSaveAction {
        yxDoNotSaveChanges = 2,
        yxSaveChanges = 1

    }
    export const enum YxSaveAsAccessMode {
        yxExclusive = 3,
        yxNoChange = 1,
        yxShared = 2

    }
    export const enum YxSaveAsAccessMode {
        yxExclusive = 3,
        yxNoChange = 1,
        yxShared = 2

    }
    export const enum YxSaveConflictResolution {
        yxLocalSessionChanges = 2,
        yxOtherSessionChanges = 3,
        yxUserResolution = 1

    }
    export const enum YxSaveConflictResolution {
        yxLocalSessionChanges = 2,
        yxOtherSessionChanges = 3,
        yxUserResolution = 1

    }
    export const enum YxScaleType {
        yxScaleLinear = -4132,
        yxScaleLogarithmic = -4133

    }
    export const enum YxScaleType {
        yxScaleLinear = -4132,
        yxScaleLogarithmic = -4133

    }
    export const enum YxSearchDirection {
        yxNext = 1,
        yxPrevious = 2

    }
    export const enum YxSearchDirection {
        yxNext = 1,
        yxPrevious = 2

    }
    export const enum YxSearchOrder {
        yxByColumns = 2,
        yxByRows = 1

    }
    export const enum YxSearchOrder {
        yxByColumns = 2,
        yxByRows = 1

    }
    export const enum YxSheetType {
        yxChart = -4109,
        yxDialogSheet = -4116,
        yxExcel4IntlMacroSheet = 4,
        yxExcel4MacroSheet = 3,
        yxWorksheet = -4167

    }
    export const enum YxSheetType {
        yxChart = -4109,
        yxDialogSheet = -4116,
        yxExcel4IntlMacroSheet = 4,
        yxExcel4MacroSheet = 3,
        yxWorksheet = -4167

    }
    export const enum YxSheetVisibility {
        yxSheetHidden = 0,
        yxSheetVeryHidden = 2,
        yxSheetVisible = -1

    }
    export const enum YxSheetVisibility {
        yxSheetHidden = 0,
        yxSheetVeryHidden = 2,
        yxSheetVisible = -1

    }
    export const enum YxSizeRepresents {
        yxSizeIsArea = 1,
        yxSizeIsWidth = 2

    }
    export const enum YxSizeRepresents {
        yxSizeIsArea = 1,
        yxSizeIsWidth = 2

    }
    export const enum YxSmartTagDisplayMode {
        yxButtonOnly = 2,
        yxDisplayNone = 1,
        yxIndicatorAndButton = 0

    }
    export const enum YxSmartTagDisplayMode {
        yxButtonOnly = 2,
        yxDisplayNone = 1,
        yxIndicatorAndButton = 0

    }
    export const enum YxSortDataOption {
        yxSortNormal = 0,
        yxSortTextAsNumbers = 1

    }
    export const enum YxSortDataOption {
        yxSortNormal = 0,
        yxSortTextAsNumbers = 1

    }
    export const enum YxSortMethod {
        yxPinYin = 1,
        yxStroke = 2

    }
    export const enum YxSortMethod {
        yxPinYin = 1,
        yxStroke = 2

    }
    export const enum YxSortOn {
        SortOnCellColor = 1,
        SortOnFontColor = 2,
        SortOnIcon = 3,
        SortOnValues = 0

    }
    export const enum YxSortOn {
        SortOnCellColor = 1,
        SortOnFontColor = 2,
        SortOnIcon = 3,
        SortOnValues = 0

    }
    export const enum YxSortOrder {
        yxAscending = 1,
        yxDescending = 2

    }
    export const enum YxSortOrder {
        yxAscending = 1,
        yxDescending = 2

    }
    export const enum YxSortOrientation {
        yxSortColumns = 1,
        yxSortRows = 2

    }
    export const enum YxSortOrientation {
        yxSortColumns = 1,
        yxSortRows = 2

    }
    export const enum YxSortType {
        yxSortLabels = 2,
        yxSortValues = 1

    }
    export const enum YxSortType {
        yxSortLabels = 2,
        yxSortValues = 1

    }
    export const enum YxSpeakDirection {
        yxSpeakByColumns = 1,
        yxSpeakByRows = 0

    }
    export const enum YxSpeakDirection {
        yxSpeakByColumns = 1,
        yxSpeakByRows = 0

    }
    export const enum YxSpecialCellsValue {
        yxErrors = 16,
        yxLogical = 4,
        yxNumbers = 1,
        yxTextValues = 2

    }
    export const enum YxSpecialCellsValue {
        yxErrors = 16,
        yxLogical = 4,
        yxNumbers = 1,
        yxTextValues = 2

    }
    export const enum YxSubscribeToFormat {
        yxSubscribeToPicture = -4147,
        yxSubscribeToText = -4158

    }
    export const enum YxSubscribeToFormat {
        yxSubscribeToPicture = -4147,
        yxSubscribeToText = -4158

    }
    export const enum YxSubtototalLocationType {
        yxAtBottom = 2,
        yxAtTop = 1

    }
    export const enum YxSubtototalLocationType {
        yxAtBottom = 2,
        yxAtTop = 1

    }
    export const enum YxSummaryReportType {
        yxStandardSummary = 1,
        yxSummaryPivotTable = -4148

    }
    export const enum YxSummaryReportType {
        yxStandardSummary = 1,
        yxSummaryPivotTable = -4148

    }
    export const enum YxSummaryRow {
        yxSummaryAbove = 0,
        yxSummaryBelow = 1

    }
    export const enum YxSummaryRow {
        yxSummaryAbove = 0,
        yxSummaryBelow = 1

    }
    export const enum YxTextParsingType {
        yxDelimited = 1,
        yxFixedWidth = 2

    }
    export const enum YxTextParsingType {
        yxDelimited = 1,
        yxFixedWidth = 2

    }
    export const enum YxTextQualifier {
        yxTextQualifierDoubleQuote = 1,
        yxTextQualifierNone = -4142,
        yxTextQualifierSingleQuote = 2

    }
    export const enum YxTextQualifier {
        yxTextQualifierDoubleQuote = 1,
        yxTextQualifierNone = -4142,
        yxTextQualifierSingleQuote = 2

    }
    export const enum YxTickLabelOrientation {
        yxTickLabelOrientationAutomatic = -4105,
        yxTickLabelOrientationDownward = -4170,
        yxTickLabelOrientationHorizontal = -4128,
        yxTickLabelOrientationUpward = -4171,
        yxTickLabelOrientationVertical = -4166

    }
    export const enum YxTickLabelOrientation {
        yxTickLabelOrientationAutomatic = -4105,
        yxTickLabelOrientationDownward = -4170,
        yxTickLabelOrientationHorizontal = -4128,
        yxTickLabelOrientationUpward = -4171,
        yxTickLabelOrientationVertical = -4166

    }
    export const enum YxTickLabelPosition {
        yxTickLabelPositionHigh = -4127,
        yxTickLabelPositionLow = -4134,
        yxTickLabelPositionNextToAxis = 4,
        yxTickLabelPositionNone = -4142

    }
    export const enum YxTickLabelPosition {
        yxTickLabelPositionHigh = -4127,
        yxTickLabelPositionLow = -4134,
        yxTickLabelPositionNextToAxis = 4,
        yxTickLabelPositionNone = -4142

    }
    export const enum YxTickMark {
        yxTickMarkCross = 4,
        yxTickMarkInside = 2,
        yxTickMarkNone = -4142,
        yxTickMarkOutside = 3

    }
    export const enum YxTickMark {
        yxTickMarkCross = 4,
        yxTickMarkInside = 2,
        yxTickMarkNone = -4142,
        yxTickMarkOutside = 3

    }
    export const enum YxTimeUnit {
        yxDays = 0,
        yxMonths = 1,
        yxYears = 2

    }
    export const enum YxTimeUnit {
        yxDays = 0,
        yxMonths = 1,
        yxYears = 2

    }
    export const enum YxTopBottom {
        xlTop10Bottom = 0,
        xlTop10Top = 1

    }
    export const enum YxTopBottom {
        xlTop10Bottom = 0,
        xlTop10Top = 1

    }
    export const enum YxTrendlineType {
        yxExponential = 5,
        yxLinear = -4132,
        yxLogarithmic = -4133,
        yxMovingAvg = 6,
        yxPolynomial = 3,
        yxPower = 4

    }
    export const enum YxTrendlineType {
        yxExponential = 5,
        yxLinear = -4132,
        yxLogarithmic = -4133,
        yxMovingAvg = 6,
        yxPolynomial = 3,
        yxPower = 4

    }
    export const enum YxUpdateLinks {
        yxUpdateLinksAlways = 3,
        yxUpdateLinksNever = 2,
        yxUpdateLinksUserSetting = 1

    }
    export const enum YxUpdateLinks {
        yxUpdateLinksAlways = 3,
        yxUpdateLinksNever = 2,
        yxUpdateLinksUserSetting = 1

    }
    export const enum YxVAlign {
        yxVAlignBottom = -4107,
        yxVAlignCenter = -4108,
        yxVAlignDistributed = -4117,
        yxVAlignJustify = -4130,
        yxVAlignTop = -4160

    }
    export const enum YxVAlign {
        yxVAlignBottom = -4107,
        yxVAlignCenter = -4108,
        yxVAlignDistributed = -4117,
        yxVAlignJustify = -4130,
        yxVAlignTop = -4160

    }
    export const enum YxWindowState {
        yxMaximized = -4137,
        yxMinimized = -4140,
        yxNormal = -4143

    }
    export const enum YxWindowState {
        yxMaximized = -4137,
        yxMinimized = -4140,
        yxNormal = -4143

    }
    export const enum YxWindowType {
        yxChartAsWindow = 5,
        yxChartInPlace = 4,
        yxClipboard = 3,
        yxInfo = -4129,
        yxWorkbook = 1

    }
    export const enum YxWindowType {
        yxChartAsWindow = 5,
        yxChartInPlace = 4,
        yxClipboard = 3,
        yxInfo = -4129,
        yxWorkbook = 1

    }
    export const enum YxWindowView {
        yxNormalView = 1,
        yxPageBreakPreview = 2

    }
    export const enum YxWindowView {
        yxNormalView = 1,
        yxPageBreakPreview = 2

    }
    export const enum YxXmlExportResult {
        yxXmlExportSuccess = 0,
        yxXmlExportValidationFailed = 1

    }
    export const enum YxXmlExportResult {
        yxXmlExportSuccess = 0,
        yxXmlExportValidationFailed = 1

    }
    export const enum YxXmlImportResult {
        yxXmlImportElementsTruncated = 1,
        yxXmlImportSuccess = 0,
        yxXmlImportValidationFailed = 2

    }
    export const enum YxXmlImportResult {
        yxXmlImportElementsTruncated = 1,
        yxXmlImportSuccess = 0,
        yxXmlImportValidationFailed = 2

    }
    export const enum YxXmlLoadOption {
        yxXmlLoadImportToList = 2,
        yxXmlLoadMapXml = 3,
        yxXmlLoadOpenXml = 1,
        yxXmlLoadPromptUser = 0

    }
    export const enum YxXmlLoadOption {
        yxXmlLoadImportToList = 2,
        yxXmlLoadMapXml = 3,
        yxXmlLoadOpenXml = 1,
        yxXmlLoadPromptUser = 0

    }
    export const enum YxYesNoGuess {
        yxGuess = 0,
        yxNo = 2,
        yxYes = 1

    }
    export const enum YxYesNoGuess {
        yxGuess = 0,
        yxNo = 2,
        yxYes = 1

    }
}
declare namespace excel.event {
    export interface ApplicationAdapter {
        WorkbookPivotTableCloseConnection(e: excel.event.ApplicationEvent): void
        sheetBeforeDoubleClick(e: excel.event.ApplicationEvent): void
        sheetBeforeRightClick(e: excel.event.ApplicationEvent): void
        sheetFollowHyperlink(e: excel.event.ApplicationEvent): void
        sheetPivotTableUpdate(e: excel.event.ApplicationEvent): void
        sheetSelectionChange(e: excel.event.ApplicationEvent): void
        WorkbookBeforePrint(e: excel.event.ApplicationEvent): void
        WorkbookBeforeSave(e: excel.event.ApplicationEvent): void
        WorkbookDeactivate(e: excel.event.ApplicationEvent): void
        WorkbookBeforeClose(e: excel.event.ApplicationEvent): void
        WorkbookAfterSave(e: excel.event.ApplicationEvent): void
        WorkbookSync(e: excel.event.ApplicationEvent): void
        sheetActivate(e: excel.event.ApplicationEvent): void
        sheetCalculate(e: excel.event.ApplicationEvent): void
        sheetChange(e: excel.event.ApplicationEvent): void
        sheetDeactivate(e: excel.event.ApplicationEvent): void
        WorkbookActivate(e: excel.event.ApplicationEvent): void
        WindowActivate(e: excel.event.ApplicationEvent): void
        WindowDeactivate(e: excel.event.ApplicationEvent): void
        WorkbookOpen(e: excel.event.ApplicationEvent): void
        newWorkbook(e: excel.event.ApplicationEvent): void
        WindowResize(e: excel.event.ApplicationEvent): void
        WorkbookNewSheet(e: excel.event.ApplicationEvent): void
        WorkbookAddinInstall(e: excel.event.ApplicationEvent): void
        WorkbookAddinUninstall(e: excel.event.ApplicationEvent): void
        WorkbookAfterXmlExport(e: excel.event.ApplicationEvent): void
        WorkbookAfterXmlImport(e: excel.event.ApplicationEvent): void
        WorkbookBeforeXmlExport(e: excel.event.ApplicationEvent): void
        WorkbookBeforeXmlImport(e: excel.event.ApplicationEvent): void
        WorkbookPivotTableOpenConnection(e: excel.event.ApplicationEvent): void
    }
    export interface ApplicationAdapter {
        WorkbookPivotTableCloseConnection(e: excel.event.ApplicationEvent): void
        sheetBeforeDoubleClick(e: excel.event.ApplicationEvent): void
        sheetBeforeRightClick(e: excel.event.ApplicationEvent): void
        sheetFollowHyperlink(e: excel.event.ApplicationEvent): void
        sheetPivotTableUpdate(e: excel.event.ApplicationEvent): void
        sheetSelectionChange(e: excel.event.ApplicationEvent): void
        WorkbookBeforePrint(e: excel.event.ApplicationEvent): void
        WorkbookBeforeSave(e: excel.event.ApplicationEvent): void
        WorkbookDeactivate(e: excel.event.ApplicationEvent): void
        WorkbookBeforeClose(e: excel.event.ApplicationEvent): void
        WorkbookAfterSave(e: excel.event.ApplicationEvent): void
        WorkbookSync(e: excel.event.ApplicationEvent): void
        sheetActivate(e: excel.event.ApplicationEvent): void
        sheetCalculate(e: excel.event.ApplicationEvent): void
        sheetChange(e: excel.event.ApplicationEvent): void
        sheetDeactivate(e: excel.event.ApplicationEvent): void
        WorkbookActivate(e: excel.event.ApplicationEvent): void
        WindowActivate(e: excel.event.ApplicationEvent): void
        WindowDeactivate(e: excel.event.ApplicationEvent): void
        WorkbookOpen(e: excel.event.ApplicationEvent): void
        newWorkbook(e: excel.event.ApplicationEvent): void
        WindowResize(e: excel.event.ApplicationEvent): void
        WorkbookNewSheet(e: excel.event.ApplicationEvent): void
        WorkbookAddinInstall(e: excel.event.ApplicationEvent): void
        WorkbookAddinUninstall(e: excel.event.ApplicationEvent): void
        WorkbookAfterXmlExport(e: excel.event.ApplicationEvent): void
        WorkbookAfterXmlImport(e: excel.event.ApplicationEvent): void
        WorkbookBeforeXmlExport(e: excel.event.ApplicationEvent): void
        WorkbookBeforeXmlImport(e: excel.event.ApplicationEvent): void
        WorkbookPivotTableOpenConnection(e: excel.event.ApplicationEvent): void
    }
    export interface ApplicationEvent {
        isCancel(): boolean
        setCancel(cancel: boolean): void
        getRange(): excel.Range
        getHyperlink(): excel.Hyperlink
        setShowSaveAsUI(show: boolean): void
        isShowSaveAsUI(): boolean
        getWindow(): excel.Window
        getWorkbook(): excel.Workbook
        getWorksheet(): excel.Worksheet
        isSuccess(): boolean
    }
    export interface ApplicationEvent {
        isCancel(): boolean
        setCancel(cancel: boolean): void
        getRange(): excel.Range
        getHyperlink(): excel.Hyperlink
        setShowSaveAsUI(show: boolean): void
        isShowSaveAsUI(): boolean
        getWindow(): excel.Window
        getWorkbook(): excel.Workbook
        getWorksheet(): excel.Worksheet
        isSuccess(): boolean
    }
    export interface ApplicationListener {
        WorkbookPivotTableCloseConnection(e: excel.event.ApplicationEvent): void
        sheetBeforeDoubleClick(e: excel.event.ApplicationEvent): void
        sheetBeforeRightClick(e: excel.event.ApplicationEvent): void
        sheetFollowHyperlink(e: excel.event.ApplicationEvent): void
        sheetPivotTableUpdate(e: excel.event.ApplicationEvent): void
        sheetSelectionChange(e: excel.event.ApplicationEvent): void
        WorkbookBeforePrint(e: excel.event.ApplicationEvent): void
        WorkbookBeforeSave(e: excel.event.ApplicationEvent): void
        WorkbookDeactivate(e: excel.event.ApplicationEvent): void
        WorkbookBeforeClose(e: excel.event.ApplicationEvent): void
        WorkbookAfterSave(e: excel.event.ApplicationEvent): void
        WorkbookSync(e: excel.event.ApplicationEvent): void
        sheetActivate(e: excel.event.ApplicationEvent): void
        sheetCalculate(e: excel.event.ApplicationEvent): void
        sheetChange(e: excel.event.ApplicationEvent): void
        sheetDeactivate(e: excel.event.ApplicationEvent): void
        WorkbookActivate(e: excel.event.ApplicationEvent): void
        WindowActivate(e: excel.event.ApplicationEvent): void
        WindowDeactivate(e: excel.event.ApplicationEvent): void
        WorkbookOpen(e: excel.event.ApplicationEvent): void
        newWorkbook(e: excel.event.ApplicationEvent): void
        WindowResize(e: excel.event.ApplicationEvent): void
        WorkbookNewSheet(e: excel.event.ApplicationEvent): void
        WorkbookAddinInstall(e: excel.event.ApplicationEvent): void
        WorkbookAddinUninstall(e: excel.event.ApplicationEvent): void
        WorkbookAfterXmlExport(e: excel.event.ApplicationEvent): void
        WorkbookAfterXmlImport(e: excel.event.ApplicationEvent): void
        WorkbookBeforeXmlExport(e: excel.event.ApplicationEvent): void
        WorkbookBeforeXmlImport(e: excel.event.ApplicationEvent): void
        WorkbookPivotTableOpenConnection(e: excel.event.ApplicationEvent): void
    }
    export interface ApplicationListener {
        WorkbookPivotTableCloseConnection(e: excel.event.ApplicationEvent): void
        sheetBeforeDoubleClick(e: excel.event.ApplicationEvent): void
        sheetBeforeRightClick(e: excel.event.ApplicationEvent): void
        sheetFollowHyperlink(e: excel.event.ApplicationEvent): void
        sheetPivotTableUpdate(e: excel.event.ApplicationEvent): void
        sheetSelectionChange(e: excel.event.ApplicationEvent): void
        WorkbookBeforePrint(e: excel.event.ApplicationEvent): void
        WorkbookBeforeSave(e: excel.event.ApplicationEvent): void
        WorkbookDeactivate(e: excel.event.ApplicationEvent): void
        WorkbookBeforeClose(e: excel.event.ApplicationEvent): void
        WorkbookAfterSave(e: excel.event.ApplicationEvent): void
        WorkbookSync(e: excel.event.ApplicationEvent): void
        sheetActivate(e: excel.event.ApplicationEvent): void
        sheetCalculate(e: excel.event.ApplicationEvent): void
        sheetChange(e: excel.event.ApplicationEvent): void
        sheetDeactivate(e: excel.event.ApplicationEvent): void
        WorkbookActivate(e: excel.event.ApplicationEvent): void
        WindowActivate(e: excel.event.ApplicationEvent): void
        WindowDeactivate(e: excel.event.ApplicationEvent): void
        WorkbookOpen(e: excel.event.ApplicationEvent): void
        newWorkbook(e: excel.event.ApplicationEvent): void
        WindowResize(e: excel.event.ApplicationEvent): void
        WorkbookNewSheet(e: excel.event.ApplicationEvent): void
        WorkbookAddinInstall(e: excel.event.ApplicationEvent): void
        WorkbookAddinUninstall(e: excel.event.ApplicationEvent): void
        WorkbookAfterXmlExport(e: excel.event.ApplicationEvent): void
        WorkbookAfterXmlImport(e: excel.event.ApplicationEvent): void
        WorkbookBeforeXmlExport(e: excel.event.ApplicationEvent): void
        WorkbookBeforeXmlImport(e: excel.event.ApplicationEvent): void
        WorkbookPivotTableOpenConnection(e: excel.event.ApplicationEvent): void
    }
    export interface MainAdapter {
        getType(): int
        open(e: excel.event.WorkbookEvent): void
        sync(e: excel.event.WorkbookEvent): void
        componentResized(e: any): void
        componentMoved(e: any): void
        componentShown(e: any): void
        componentHidden(e: any): void
        focusGained(e: any): void
        focusLost(e: any): void
        windowOpened(e: any): void
        windowClosing(e: any): void
        windowClosed(e: any): void
        windowIconified(e: any): void
        windowDeiconified(e: any): void
        windowActivated(e: any): void
        windowDeactivated(e: any): void
        activate(e: excel.event.WorkbookEvent): void
        activate(e: excel.event.WorksheetEvent): void
        change(e: excel.event.WorksheetEvent): void
        statusChanged(event: any): boolean
        windowActivate(e: excel.event.WorkbookEvent): void
        windowActivate(event: any): void
        beforeExit(event: any): boolean
        saveTree(event: any): void
        loadTree(event: any): void
        windowDeactivate(e: excel.event.WorkbookEvent): void
        windowDeactivate(event: any): void
        windowResize(e: excel.event.WorkbookEvent): void
        windowResize(event: any): void
        afterOpenBinder(event: any): void
        workbookActivate(event: any): void
        saveTemplate(event: any): void
        loadTemplate(event: any): void
        delTemplate(event: any): void
        loadCustomMeta(event: any): void
        loadFixedBase(event: any): void
        wordStatusChange(event: any): boolean
        beforeCloseBinder(event: any): boolean
        afterCreateBinder(event: any): void
        beforePrintWorkbook(event: any): boolean
        afterPrintWorkbook(event: any): boolean
        beforeSaveWorkbook(event: any): boolean
        workbookDeactivate(event: any): void
        afterBinderReveredSave(event: any): void
        followHyperlink(e: excel.event.WorksheetEvent): void
        calculate(e: excel.event.WorksheetEvent): void
        newSheet(e: excel.event.WorkbookEvent): void
        pivotTableCloseConnection(e: excel.event.WorkbookEvent): void
        pivotTableOpenConnection(e: excel.event.WorkbookEvent): void
        sheetBeforeDoubleClick(e: excel.event.WorkbookEvent): void
        sheetBeforeRightClick(e: excel.event.WorkbookEvent): void
        sheetFollowHyperlink(e: excel.event.WorkbookEvent): void
        sheetPivotTableUpdate(e: excel.event.WorkbookEvent): void
        sheetSelectionChange(e: excel.event.WorkbookEvent): void
        beforeDoubleClick(e: excel.event.WorksheetEvent): void
        addinInstall(e: excel.event.WorkbookEvent): void
        addinUninstall(e: excel.event.WorkbookEvent): void
        afterXmlExport(e: excel.event.WorkbookEvent): void
        afterXmlImport(e: excel.event.WorkbookEvent): void
        beforeClose(e: excel.event.WorkbookEvent): void
        beforePrint(e: excel.event.WorkbookEvent): void
        beforeSave(e: excel.event.WorkbookEvent): void
        afterSave(e: excel.event.WorkbookEvent): void
        beforeXmlExport(e: excel.event.WorkbookEvent): void
        beforeXmlImport(e: excel.event.WorkbookEvent): void
        deactivate(e: excel.event.WorksheetEvent): void
        deactivate(e: excel.event.WorkbookEvent): void
        sheetActivate(e: excel.event.WorkbookEvent): void
        sheetCalculate(e: excel.event.WorkbookEvent): void
        sheetChange(e: excel.event.WorkbookEvent): void
        sheetDeactivate(e: excel.event.WorkbookEvent): void
        beforeRightClick(e: excel.event.WorksheetEvent): void
        pivotTableUpdate(e: excel.event.WorksheetEvent): void
        selectionChange(e: excel.event.WorksheetEvent): void
    }
    export interface MainAdapter {
        getType(): int
        open(e: excel.event.WorkbookEvent): void
        sync(e: excel.event.WorkbookEvent): void
        componentResized(e: any): void
        componentMoved(e: any): void
        componentShown(e: any): void
        componentHidden(e: any): void
        focusGained(e: any): void
        focusLost(e: any): void
        windowOpened(e: any): void
        windowClosing(e: any): void
        windowClosed(e: any): void
        windowIconified(e: any): void
        windowDeiconified(e: any): void
        windowActivated(e: any): void
        windowDeactivated(e: any): void
        activate(e: excel.event.WorkbookEvent): void
        activate(e: excel.event.WorksheetEvent): void
        change(e: excel.event.WorksheetEvent): void
        statusChanged(event: any): boolean
        windowActivate(e: excel.event.WorkbookEvent): void
        windowActivate(event: any): void
        beforeExit(event: any): boolean
        saveTree(event: any): void
        loadTree(event: any): void
        windowDeactivate(e: excel.event.WorkbookEvent): void
        windowDeactivate(event: any): void
        windowResize(e: excel.event.WorkbookEvent): void
        windowResize(event: any): void
        afterOpenBinder(event: any): void
        workbookActivate(event: any): void
        saveTemplate(event: any): void
        loadTemplate(event: any): void
        delTemplate(event: any): void
        loadCustomMeta(event: any): void
        loadFixedBase(event: any): void
        wordStatusChange(event: any): boolean
        beforeCloseBinder(event: any): boolean
        afterCreateBinder(event: any): void
        beforePrintWorkbook(event: any): boolean
        afterPrintWorkbook(event: any): boolean
        beforeSaveWorkbook(event: any): boolean
        workbookDeactivate(event: any): void
        afterBinderReveredSave(event: any): void
        followHyperlink(e: excel.event.WorksheetEvent): void
        calculate(e: excel.event.WorksheetEvent): void
        newSheet(e: excel.event.WorkbookEvent): void
        pivotTableCloseConnection(e: excel.event.WorkbookEvent): void
        pivotTableOpenConnection(e: excel.event.WorkbookEvent): void
        sheetBeforeDoubleClick(e: excel.event.WorkbookEvent): void
        sheetBeforeRightClick(e: excel.event.WorkbookEvent): void
        sheetFollowHyperlink(e: excel.event.WorkbookEvent): void
        sheetPivotTableUpdate(e: excel.event.WorkbookEvent): void
        sheetSelectionChange(e: excel.event.WorkbookEvent): void
        beforeDoubleClick(e: excel.event.WorksheetEvent): void
        addinInstall(e: excel.event.WorkbookEvent): void
        addinUninstall(e: excel.event.WorkbookEvent): void
        afterXmlExport(e: excel.event.WorkbookEvent): void
        afterXmlImport(e: excel.event.WorkbookEvent): void
        beforeClose(e: excel.event.WorkbookEvent): void
        beforePrint(e: excel.event.WorkbookEvent): void
        beforeSave(e: excel.event.WorkbookEvent): void
        afterSave(e: excel.event.WorkbookEvent): void
        beforeXmlExport(e: excel.event.WorkbookEvent): void
        beforeXmlImport(e: excel.event.WorkbookEvent): void
        deactivate(e: excel.event.WorksheetEvent): void
        deactivate(e: excel.event.WorkbookEvent): void
        sheetActivate(e: excel.event.WorkbookEvent): void
        sheetCalculate(e: excel.event.WorkbookEvent): void
        sheetChange(e: excel.event.WorkbookEvent): void
        sheetDeactivate(e: excel.event.WorkbookEvent): void
        beforeRightClick(e: excel.event.WorksheetEvent): void
        pivotTableUpdate(e: excel.event.WorksheetEvent): void
        selectionChange(e: excel.event.WorksheetEvent): void
    }
    export const enum OLEObjectEvent {
        //serialVersionUID=
    }
    export const enum OLEObjectEvent {
        //serialVersionUID=
    }
    export interface OLEObjectListener {
        gotFocus(e: excel.event.OLEObjectEvent): void
        lostFocus(e: excel.event.OLEObjectEvent): void
    }
    export interface OLEObjectListener {
        gotFocus(e: excel.event.OLEObjectEvent): void
        lostFocus(e: excel.event.OLEObjectEvent): void
    }
    export const enum QueryTableEvent {
        //serialVersionUID=
    }
    export const enum QueryTableEvent {
        //serialVersionUID=
    }
    export interface QueryTableListener {
        afterRefresh(e: excel.event.QueryTableEvent, success: boolean): void
        beforeRefresh(e: excel.event.QueryTableEvent, cancel: boolean): void
    }
    export interface QueryTableListener {
        afterRefresh(e: excel.event.QueryTableEvent, success: boolean): void
        beforeRefresh(e: excel.event.QueryTableEvent, cancel: boolean): void
    }
    export interface WorkbookAdapter {
        open(e: excel.event.WorkbookEvent): void
        sync(e: excel.event.WorkbookEvent): void
        activate(e: excel.event.WorkbookEvent): void
        windowActivate(e: excel.event.WorkbookEvent): void
        windowDeactivate(e: excel.event.WorkbookEvent): void
        windowResize(e: excel.event.WorkbookEvent): void
        newSheet(e: excel.event.WorkbookEvent): void
        pivotTableCloseConnection(e: excel.event.WorkbookEvent): void
        pivotTableOpenConnection(e: excel.event.WorkbookEvent): void
        sheetBeforeDoubleClick(e: excel.event.WorkbookEvent): void
        sheetBeforeRightClick(e: excel.event.WorkbookEvent): void
        sheetFollowHyperlink(e: excel.event.WorkbookEvent): void
        sheetPivotTableUpdate(e: excel.event.WorkbookEvent): void
        sheetSelectionChange(e: excel.event.WorkbookEvent): void
        addinInstall(e: excel.event.WorkbookEvent): void
        addinUninstall(e: excel.event.WorkbookEvent): void
        afterXmlExport(e: excel.event.WorkbookEvent): void
        afterXmlImport(e: excel.event.WorkbookEvent): void
        beforeClose(e: excel.event.WorkbookEvent): void
        beforePrint(e: excel.event.WorkbookEvent): void
        beforeSave(e: excel.event.WorkbookEvent): void
        afterSave(e: excel.event.WorkbookEvent): void
        beforeXmlExport(e: excel.event.WorkbookEvent): void
        beforeXmlImport(e: excel.event.WorkbookEvent): void
        deactivate(e: excel.event.WorkbookEvent): void
        sheetActivate(e: excel.event.WorkbookEvent): void
        sheetCalculate(e: excel.event.WorkbookEvent): void
        sheetChange(e: excel.event.WorkbookEvent): void
        sheetDeactivate(e: excel.event.WorkbookEvent): void
    }
    export interface WorkbookAdapter {
        open(e: excel.event.WorkbookEvent): void
        sync(e: excel.event.WorkbookEvent): void
        activate(e: excel.event.WorkbookEvent): void
        windowActivate(e: excel.event.WorkbookEvent): void
        windowDeactivate(e: excel.event.WorkbookEvent): void
        windowResize(e: excel.event.WorkbookEvent): void
        newSheet(e: excel.event.WorkbookEvent): void
        pivotTableCloseConnection(e: excel.event.WorkbookEvent): void
        pivotTableOpenConnection(e: excel.event.WorkbookEvent): void
        sheetBeforeDoubleClick(e: excel.event.WorkbookEvent): void
        sheetBeforeRightClick(e: excel.event.WorkbookEvent): void
        sheetFollowHyperlink(e: excel.event.WorkbookEvent): void
        sheetPivotTableUpdate(e: excel.event.WorkbookEvent): void
        sheetSelectionChange(e: excel.event.WorkbookEvent): void
        addinInstall(e: excel.event.WorkbookEvent): void
        addinUninstall(e: excel.event.WorkbookEvent): void
        afterXmlExport(e: excel.event.WorkbookEvent): void
        afterXmlImport(e: excel.event.WorkbookEvent): void
        beforeClose(e: excel.event.WorkbookEvent): void
        beforePrint(e: excel.event.WorkbookEvent): void
        beforeSave(e: excel.event.WorkbookEvent): void
        afterSave(e: excel.event.WorkbookEvent): void
        beforeXmlExport(e: excel.event.WorkbookEvent): void
        beforeXmlImport(e: excel.event.WorkbookEvent): void
        deactivate(e: excel.event.WorkbookEvent): void
        sheetActivate(e: excel.event.WorkbookEvent): void
        sheetCalculate(e: excel.event.WorkbookEvent): void
        sheetChange(e: excel.event.WorkbookEvent): void
        sheetDeactivate(e: excel.event.WorkbookEvent): void
    }
    export interface WorkbookEvent {
        getType(): int
        setType(type: int): void
        isCancel(): boolean
        setCancel(cancel: boolean): void
        getRange(): excel.Range
        getHyperlink(): excel.Hyperlink
        setShowSaveAsUI(show: boolean): void
        isShowSaveAsUI(): boolean
        getWindow(): excel.Window
        getWorkbook(): excel.Workbook
        getWorksheet(): excel.Worksheet
        isSuccess(): boolean
    }
    export interface WorkbookEvent {
        getType(): int
        setType(type: int): void
        isCancel(): boolean
        setCancel(cancel: boolean): void
        getRange(): excel.Range
        getHyperlink(): excel.Hyperlink
        setShowSaveAsUI(show: boolean): void
        isShowSaveAsUI(): boolean
        getWindow(): excel.Window
        getWorkbook(): excel.Workbook
        getWorksheet(): excel.Worksheet
        isSuccess(): boolean
    }
    export interface WorkbookListener {
        open(e: excel.event.WorkbookEvent): void
        sync(e: excel.event.WorkbookEvent): void
        activate(e: excel.event.WorkbookEvent): void
        windowActivate(e: excel.event.WorkbookEvent): void
        windowDeactivate(e: excel.event.WorkbookEvent): void
        windowResize(e: excel.event.WorkbookEvent): void
        newSheet(e: excel.event.WorkbookEvent): void
        pivotTableCloseConnection(e: excel.event.WorkbookEvent): void
        pivotTableOpenConnection(e: excel.event.WorkbookEvent): void
        sheetBeforeDoubleClick(e: excel.event.WorkbookEvent): void
        sheetBeforeRightClick(e: excel.event.WorkbookEvent): void
        sheetFollowHyperlink(e: excel.event.WorkbookEvent): void
        sheetPivotTableUpdate(e: excel.event.WorkbookEvent): void
        sheetSelectionChange(e: excel.event.WorkbookEvent): void
        addinInstall(e: excel.event.WorkbookEvent): void
        addinUninstall(e: excel.event.WorkbookEvent): void
        afterXmlExport(e: excel.event.WorkbookEvent): void
        afterXmlImport(e: excel.event.WorkbookEvent): void
        beforeClose(e: excel.event.WorkbookEvent): void
        beforePrint(e: excel.event.WorkbookEvent): void
        beforeSave(e: excel.event.WorkbookEvent): void
        afterSave(e: excel.event.WorkbookEvent): void
        beforeXmlExport(e: excel.event.WorkbookEvent): void
        beforeXmlImport(e: excel.event.WorkbookEvent): void
        deactivate(e: excel.event.WorkbookEvent): void
        sheetActivate(e: excel.event.WorkbookEvent): void
        sheetCalculate(e: excel.event.WorkbookEvent): void
        sheetChange(e: excel.event.WorkbookEvent): void
        sheetDeactivate(e: excel.event.WorkbookEvent): void
    }
    export interface WorkbookListener {
        open(e: excel.event.WorkbookEvent): void
        sync(e: excel.event.WorkbookEvent): void
        activate(e: excel.event.WorkbookEvent): void
        windowActivate(e: excel.event.WorkbookEvent): void
        windowDeactivate(e: excel.event.WorkbookEvent): void
        windowResize(e: excel.event.WorkbookEvent): void
        newSheet(e: excel.event.WorkbookEvent): void
        pivotTableCloseConnection(e: excel.event.WorkbookEvent): void
        pivotTableOpenConnection(e: excel.event.WorkbookEvent): void
        sheetBeforeDoubleClick(e: excel.event.WorkbookEvent): void
        sheetBeforeRightClick(e: excel.event.WorkbookEvent): void
        sheetFollowHyperlink(e: excel.event.WorkbookEvent): void
        sheetPivotTableUpdate(e: excel.event.WorkbookEvent): void
        sheetSelectionChange(e: excel.event.WorkbookEvent): void
        addinInstall(e: excel.event.WorkbookEvent): void
        addinUninstall(e: excel.event.WorkbookEvent): void
        afterXmlExport(e: excel.event.WorkbookEvent): void
        afterXmlImport(e: excel.event.WorkbookEvent): void
        beforeClose(e: excel.event.WorkbookEvent): void
        beforePrint(e: excel.event.WorkbookEvent): void
        beforeSave(e: excel.event.WorkbookEvent): void
        afterSave(e: excel.event.WorkbookEvent): void
        beforeXmlExport(e: excel.event.WorkbookEvent): void
        beforeXmlImport(e: excel.event.WorkbookEvent): void
        deactivate(e: excel.event.WorkbookEvent): void
        sheetActivate(e: excel.event.WorkbookEvent): void
        sheetCalculate(e: excel.event.WorkbookEvent): void
        sheetChange(e: excel.event.WorkbookEvent): void
        sheetDeactivate(e: excel.event.WorkbookEvent): void
    }
    export interface WorksheetAdapter {
        activate(e: excel.event.WorksheetEvent): void
        change(e: excel.event.WorksheetEvent): void
        followHyperlink(e: excel.event.WorksheetEvent): void
        calculate(e: excel.event.WorksheetEvent): void
        beforeDoubleClick(e: excel.event.WorksheetEvent): void
        deactivate(e: excel.event.WorksheetEvent): void
        beforeRightClick(e: excel.event.WorksheetEvent): void
        pivotTableUpdate(e: excel.event.WorksheetEvent): void
        selectionChange(e: excel.event.WorksheetEvent): void
    }
    export interface WorksheetAdapter {
        activate(e: excel.event.WorksheetEvent): void
        change(e: excel.event.WorksheetEvent): void
        followHyperlink(e: excel.event.WorksheetEvent): void
        calculate(e: excel.event.WorksheetEvent): void
        beforeDoubleClick(e: excel.event.WorksheetEvent): void
        deactivate(e: excel.event.WorksheetEvent): void
        beforeRightClick(e: excel.event.WorksheetEvent): void
        pivotTableUpdate(e: excel.event.WorksheetEvent): void
        selectionChange(e: excel.event.WorksheetEvent): void
    }
    export interface WorksheetEvent {
        getAddress(): string
        getType(): int
        isCancel(): boolean
        setCancel(cancel: boolean): void
        getRange(): excel.Range
        getHyperlink(): excel.Hyperlink
        getWorkbook(): excel.Workbook
        getWorksheet(): excel.Worksheet
    }
    export interface WorksheetEvent {
        getAddress(): string
        getType(): int
        isCancel(): boolean
        setCancel(cancel: boolean): void
        getRange(): excel.Range
        getHyperlink(): excel.Hyperlink
        getWorkbook(): excel.Workbook
        getWorksheet(): excel.Worksheet
    }
    export interface WorksheetListener {
        activate(e: excel.event.WorksheetEvent): void
        change(e: excel.event.WorksheetEvent): void
        followHyperlink(e: excel.event.WorksheetEvent): void
        calculate(e: excel.event.WorksheetEvent): void
        beforeDoubleClick(e: excel.event.WorksheetEvent): void
        deactivate(e: excel.event.WorksheetEvent): void
        beforeRightClick(e: excel.event.WorksheetEvent): void
        pivotTableUpdate(e: excel.event.WorksheetEvent): void
        selectionChange(e: excel.event.WorksheetEvent): void
    }
    export interface WorksheetListener {
        activate(e: excel.event.WorksheetEvent): void
        change(e: excel.event.WorksheetEvent): void
        followHyperlink(e: excel.event.WorksheetEvent): void
        calculate(e: excel.event.WorksheetEvent): void
        beforeDoubleClick(e: excel.event.WorksheetEvent): void
        deactivate(e: excel.event.WorksheetEvent): void
        beforeRightClick(e: excel.event.WorksheetEvent): void
        pivotTableUpdate(e: excel.event.WorksheetEvent): void
        selectionChange(e: excel.event.WorksheetEvent): void
    }
}
declare namespace excel.resource {
    export const enum ChartExceptionConstants {
        //AXIS_TYPE_ERROR=请使用 FIRST_AXIS/SECOND_AXIS 常量来表示第几根座标轴！,
        //INVALID_NAME=该名称无效,
        //EXISTING_SHEET=工作表已存在,
        //CANNOT_ADD_SHEET=工作表数已达最大数,
        //CHART_NOT_EXISTING=该工作表不存在,
        //TYPE_NOT_EXIST=该图表类型下没有次 Y 轴,
        //CHART_NO_EXISTING=图表不存在

    }
    export const enum ChartExceptionConstants {
        //AXIS_TYPE_ERROR=请使用 FIRST_AXIS/SECOND_AXIS 常量来表示第几根座标轴！,
        //INVALID_NAME=该名称无效,
        //EXISTING_SHEET=工作表已存在,
        //CANNOT_ADD_SHEET=工作表数已达最大数,
        //CHART_NOT_EXISTING=该工作表不存在,
        //TYPE_NOT_EXIST=该图表类型下没有次 Y 轴,
        //CHART_NO_EXISTING=图表不存在

    }
    export const enum ComExceptionConstants {
        //NO_WINDOW=打开窗口失败,
        //NO_SELECTED_OBJECT=没有对象被选中,必须先选中对象,
        //SHAPE_TYPE_ERROR=自选图形类型错误,
        //CANNOT_EDIT=此对象不(全)能编辑,
        //NO_SHAPE=选择的图形对象不存在,
        //CANNOT_MAKE_ARROW=此对象不能(全)有箭头,
        //NO_DOCFIELDNAME=公文域中含有非法字符,或名字与现有公文域名字重复,
        //ERROR_DOCFIELD=公文域在文字处理中有效,
        //LOCKED_DOCFIELD=公文域内容已锁定,
        //LOCKED_POSITION_DOCFIELD=公文域位置已锁定,
        //NO_OBJECT=没有该对象,
        //NO_Seal_File=读文件或者密码错误,
        //NO_Seal=此对象不是签章,
        //Seal_Sign=此对象已经数字签名了,
        //NOT_IMAGE=此对象不(全)是图片,
        //NOT_CONNECT=此对象不(全)是连接线,
        //DELETE_DOCFIELD=此公文域已设为不能删除,
        //NOT_PRESENTATION=当前不在演示文稿中,配色方案无意义,
        //NO_ARROW_STYLE=无此类型箭头风格,
        //NOT_TEXT_EFFECT=此对象不(全)是艺术字,
        //NO_WORKBOOK=该工作簿还没有打开,
        //NO_WORKBOOK_TYPE=该工作簿类型参数无效,
        //NO_CONNECTOR_TYPE=该连接线类型参数错误,
        //NO_BUILDER=必须先调用 buildFreeform() 方法构建曲线,
        //WRONG_FREEFORM_TYPE=此曲线类型参数错误,
        //Repeat_Name=名字已经存在，请重新输入新的名字,
        //NO_PRODUCT=产品类型参数超出可以接受的范围,
        //NO_UNDO_TIMES=撤消次数必须介于 5 和 64 之间，请重新输入本区间的数字,
        //NO_PATH_NAME=当前要插入的图片文件不存在，请重新输入图片路径,
        //DELETE_SHAPE=此自选图形已设为不能删除,
        //ERROR_MACRO_NAME=宏名称不规范,
        //NO_MACRO_FOUND=无法找到指定的宏,
        //NOT_SUPPORT_DIALOG_SHOW=当前状态下不能弹出此对话框,
        //NOT_SUPPORT_DIALOG_TYPE=对话框类型错误,
        //FILE_NOT_FIND=无法找到文件,
        //DIRECTRY_NOT_FIND=无法找到要转换的文档目录,
        //NO_TOOLBAR=此工具栏没有显示,
        //NO_PRINT_PREVIEW=在打印预览时不能做此操作。请先退出打印预览后重试。,
        //NO_BOOK_NAME=当前永中Office 中已有打开的同名文件。请先关闭当前同名文件。,
        //NO_PRINT=当前文档已禁止打印，请先关闭禁止打印选项,
        //NO_PARAMETER=输入参数不符合范围，请重新输入,
        //NAME_TOO_LONG=名称字符串过长，长度请不要超过 1024,
        //NAME_NOT_VALID=文件名称不合法，或文件路径不存在

    }
    export const enum ComExceptionConstants {
        //NO_WINDOW=打开窗口失败,
        //NO_SELECTED_OBJECT=没有对象被选中,必须先选中对象,
        //SHAPE_TYPE_ERROR=自选图形类型错误,
        //CANNOT_EDIT=此对象不(全)能编辑,
        //NO_SHAPE=选择的图形对象不存在,
        //CANNOT_MAKE_ARROW=此对象不能(全)有箭头,
        //NO_DOCFIELDNAME=公文域中含有非法字符,或名字与现有公文域名字重复,
        //ERROR_DOCFIELD=公文域在文字处理中有效,
        //LOCKED_DOCFIELD=公文域内容已锁定,
        //LOCKED_POSITION_DOCFIELD=公文域位置已锁定,
        //NO_OBJECT=没有该对象,
        //NO_Seal_File=读文件或者密码错误,
        //NO_Seal=此对象不是签章,
        //Seal_Sign=此对象已经数字签名了,
        //NOT_IMAGE=此对象不(全)是图片,
        //NOT_CONNECT=此对象不(全)是连接线,
        //DELETE_DOCFIELD=此公文域已设为不能删除,
        //NOT_PRESENTATION=当前不在演示文稿中,配色方案无意义,
        //NO_ARROW_STYLE=无此类型箭头风格,
        //NOT_TEXT_EFFECT=此对象不(全)是艺术字,
        //NO_WORKBOOK=该工作簿还没有打开,
        //NO_WORKBOOK_TYPE=该工作簿类型参数无效,
        //NO_CONNECTOR_TYPE=该连接线类型参数错误,
        //NO_BUILDER=必须先调用 buildFreeform() 方法构建曲线,
        //WRONG_FREEFORM_TYPE=此曲线类型参数错误,
        //Repeat_Name=名字已经存在，请重新输入新的名字,
        //NO_PRODUCT=产品类型参数超出可以接受的范围,
        //NO_UNDO_TIMES=撤消次数必须介于 5 和 64 之间，请重新输入本区间的数字,
        //NO_PATH_NAME=当前要插入的图片文件不存在，请重新输入图片路径,
        //DELETE_SHAPE=此自选图形已设为不能删除,
        //ERROR_MACRO_NAME=宏名称不规范,
        //NO_MACRO_FOUND=无法找到指定的宏,
        //NOT_SUPPORT_DIALOG_SHOW=当前状态下不能弹出此对话框,
        //NOT_SUPPORT_DIALOG_TYPE=对话框类型错误,
        //FILE_NOT_FIND=无法找到文件,
        //DIRECTRY_NOT_FIND=无法找到要转换的文档目录,
        //NO_TOOLBAR=此工具栏没有显示,
        //NO_PRINT_PREVIEW=在打印预览时不能做此操作。请先退出打印预览后重试。,
        //NO_BOOK_NAME=当前永中Office 中已有打开的同名文件。请先关闭当前同名文件。,
        //NO_PRINT=当前文档已禁止打印，请先关闭禁止打印选项,
        //NO_PARAMETER=输入参数不符合范围，请重新输入,
        //NAME_TOO_LONG=名称字符串过长，长度请不要超过 1024,
        //NAME_NOT_VALID=文件名称不合法，或文件路径不存在

    }
    export const enum ImportExternDataConstants {

    }
    export const enum ImportExternDataConstants {

    }
    export const enum SsExceptionConstants {
        //WORKBOOK_NOT_EXSIT=文件不存在,
        //WORKSHEET_NOT_EXSIT=工作表不存在: ,
        //NAME_ERROR=请输入正确的页面格式名称,
        //DELETE_NAME_ERROR=不能删除该页面格式,
        //RANGEADDRESS_ERROR=给定的区域地址错误或越界: ,
        //ZOOM_SIZE_ERROR=缩放比例必须在 10 到 500 之间,请输入一个该范围内的数值,
        //BETWEEN=之间,
        //DATA_SOURCE_ERROR=没有选择数据区域,
        //VALUE_ERROR=数值越界: ,
        //FIX_ERROR=修订状态,不能操作,
        //FUNCTION_MORE_RANGE_ERROR=区域中的参数太多,
        //FUNCTION_TYPE_MATCH_ERROR=参数只能是单个单元格引用,
        //FUNCTION_VALUE_ERROR=*VALUE!,
        //FUNCTION_DIV_ERROR=*DIV/0!,
        //FUNCTION_N_A_ERROR=*N/A,
        //FUNCTION_NULL_ERROR=*NULL!,
        //FUNCTION_NAME_ERROR=*NAME?,
        //FUNCTION_NUM_ERROR=*NUM!,
        //FUNCTION_REF_ERROR=*REF!,
        //TIME_SERIES_NAME_ERROR=字段名与时间序列中的现有字段名相同,
        //TIME_SERIES_NAME_NOT_EXIST=字段名不存在,
        //SCENARIO_EXIST=方案已经存在,
        //SCENARIO_NOT_EXIST=方案不存在,
        //SCENARIO_NAME_EMPTY=方案名不能为空,
        //MAX_SCENARIO=一个工作表最多只能有 252 个方案,
        //TOO_MANY_CELLS=可变单元格太多,一份方案最多只能有三十二个可变单元格,
        //MERGE_IN_SAME_WORKSHEET=无法合并同一工作表内的方案,
        //REFERENCE_ERROR=引用无效,
        //SPREADSHEET_NOT_OPEN=当前电子表格应用程序未被激活,
        //SPREADSHEET_DELETE_ERROR=集成文件中必须至少包含一个可见的工作表,
        //WORKSHEET_EXIST=工作表已经存在: ,
        //INDEX_ERROR=工作表插入位置必须大于 0,
        //ROW_LINE_NUMBER_ERROR=行号必须在 1 到 65536: ,
        //COLUMN_LINE_NUMBER_ERROR=列号必须在 1 到 256: ,
        //INVALIDATE_OPERATION=无效操作,
        //DECIMAL_ERROR=随机数的小数位取值范围为 0 到 15,
        //RENAME_ERROR=工作表重命名无效，其原因可能是： 
        //﹡工作表名称中含有非法字符。
        //﹡工作表名称与现有的某一工作表名称相同。
        //﹡工作表名称为空。,
        //KEY_WORD_ERROR=“日志”为保留名称，请重新输入其他名称。,
        //CHART_SHEET_ERROR=此操作不适用于图表工作表,
        //NAME_IS_NULL=工作表名称不能为空: ,
        //RANGE_ERROR=区域的起始行或列必须小于等于结束行或列: ,
        //ROW_HEIGHT_ERROR=工作表行高在 0 至 155.57 毫米之间: ,
        //COLUMN_WIDTH_ERROR=工作表列宽在 0 至 539.75 毫米之间: ,
        //PROTECT_ERROR=工作表单元格已被保护 ,
        //CELL_ERROR=不是单元格地址: ,
        //PATH_ERROR=文件名路径错误: ,
        //SPLIT_ERROR=工作表已有拆分,
        //FREEZE_PANE_ERROR=工作表已有冻结窗格,
        //ROW_HEIGHT_ERROR2=行高必须在 0 到 ,
        //COLUMN_WIDTH_ERROR2=列宽必须在 0 到 ,
        //NAMES_ERROR=该名称无效,
        //CROSS_ERROR=单元格区域交叉或包含: ,
        //PRINT_COMMENT_NUMBER=该数字必须介于 1 和 32767 之间:,
        //NOT_ALLOWED_EDIT=该单元格区域不允许编辑,
        //CANNOT_INSERT=为防止数据丢失，不能把非空白单元格移出集成文件,
        //STYLE_NORMAL_REMOVE=常规样式名不可删除,
        //ROW_ERROR=请选择至少包含两行单元格区域,
        //EXIST=存在,
        //NONE=不,
        //VIEW=视图,
        //POINT=磅,
        //MILLIMETER=毫米,
        //ALREADY=已,
        //STRUCTURE_PROTECTED=工作簿结构受保护,
        //WINDOWS_PROTECTED=工作簿窗口受保护,
        //PRINTER_CAN_NOT_CONNECT=打印机无法连接！,
        //PRINTER_NOT_EXIST=不存在打印机！,
        //TITLE_EXISTED=标题的区域已经存在！,
        //TITLE_HAS_ERROR_CHARACTER=含有非法字符,区域标题只能包含字母、数字和空格，并且必须是以字母开头,
        //TITLE=标题,
        //PROTECT_SHEET_ERROR=工作表已被保护

    }
    export const enum SsExceptionConstants {
        //WORKBOOK_NOT_EXSIT=文件不存在,
        //WORKSHEET_NOT_EXSIT=工作表不存在: ,
        //NAME_ERROR=请输入正确的页面格式名称,
        //DELETE_NAME_ERROR=不能删除该页面格式,
        //RANGEADDRESS_ERROR=给定的区域地址错误或越界: ,
        //ZOOM_SIZE_ERROR=缩放比例必须在 10 到 500 之间,请输入一个该范围内的数值,
        //BETWEEN=之间,
        //DATA_SOURCE_ERROR=没有选择数据区域,
        //VALUE_ERROR=数值越界: ,
        //FIX_ERROR=修订状态,不能操作,
        //FUNCTION_MORE_RANGE_ERROR=区域中的参数太多,
        //FUNCTION_TYPE_MATCH_ERROR=参数只能是单个单元格引用,
        //FUNCTION_VALUE_ERROR=*VALUE!,
        //FUNCTION_DIV_ERROR=*DIV/0!,
        //FUNCTION_N_A_ERROR=*N/A,
        //FUNCTION_NULL_ERROR=*NULL!,
        //FUNCTION_NAME_ERROR=*NAME?,
        //FUNCTION_NUM_ERROR=*NUM!,
        //FUNCTION_REF_ERROR=*REF!,
        //TIME_SERIES_NAME_ERROR=字段名与时间序列中的现有字段名相同,
        //TIME_SERIES_NAME_NOT_EXIST=字段名不存在,
        //SCENARIO_EXIST=方案已经存在,
        //SCENARIO_NOT_EXIST=方案不存在,
        //SCENARIO_NAME_EMPTY=方案名不能为空,
        //MAX_SCENARIO=一个工作表最多只能有 252 个方案,
        //TOO_MANY_CELLS=可变单元格太多,一份方案最多只能有三十二个可变单元格,
        //MERGE_IN_SAME_WORKSHEET=无法合并同一工作表内的方案,
        //REFERENCE_ERROR=引用无效,
        //SPREADSHEET_NOT_OPEN=当前电子表格应用程序未被激活,
        //SPREADSHEET_DELETE_ERROR=集成文件中必须至少包含一个可见的工作表,
        //WORKSHEET_EXIST=工作表已经存在: ,
        //INDEX_ERROR=工作表插入位置必须大于 0,
        //ROW_LINE_NUMBER_ERROR=行号必须在 1 到 65536: ,
        //COLUMN_LINE_NUMBER_ERROR=列号必须在 1 到 256: ,
        //INVALIDATE_OPERATION=无效操作,
        //DECIMAL_ERROR=随机数的小数位取值范围为 0 到 15,
        //RENAME_ERROR=工作表重命名无效，其原因可能是： 
        //﹡工作表名称中含有非法字符。
        //﹡工作表名称与现有的某一工作表名称相同。
        //﹡工作表名称为空。,
        //KEY_WORD_ERROR=“日志”为保留名称，请重新输入其他名称。,
        //CHART_SHEET_ERROR=此操作不适用于图表工作表,
        //NAME_IS_NULL=工作表名称不能为空: ,
        //RANGE_ERROR=区域的起始行或列必须小于等于结束行或列: ,
        //ROW_HEIGHT_ERROR=工作表行高在 0 至 155.57 毫米之间: ,
        //COLUMN_WIDTH_ERROR=工作表列宽在 0 至 539.75 毫米之间: ,
        //PROTECT_ERROR=工作表单元格已被保护 ,
        //CELL_ERROR=不是单元格地址: ,
        //PATH_ERROR=文件名路径错误: ,
        //SPLIT_ERROR=工作表已有拆分,
        //FREEZE_PANE_ERROR=工作表已有冻结窗格,
        //ROW_HEIGHT_ERROR2=行高必须在 0 到 ,
        //COLUMN_WIDTH_ERROR2=列宽必须在 0 到 ,
        //NAMES_ERROR=该名称无效,
        //CROSS_ERROR=单元格区域交叉或包含: ,
        //PRINT_COMMENT_NUMBER=该数字必须介于 1 和 32767 之间:,
        //NOT_ALLOWED_EDIT=该单元格区域不允许编辑,
        //CANNOT_INSERT=为防止数据丢失，不能把非空白单元格移出集成文件,
        //STYLE_NORMAL_REMOVE=常规样式名不可删除,
        //ROW_ERROR=请选择至少包含两行单元格区域,
        //EXIST=存在,
        //NONE=不,
        //VIEW=视图,
        //POINT=磅,
        //MILLIMETER=毫米,
        //ALREADY=已,
        //STRUCTURE_PROTECTED=工作簿结构受保护,
        //WINDOWS_PROTECTED=工作簿窗口受保护,
        //PRINTER_CAN_NOT_CONNECT=打印机无法连接！,
        //PRINTER_NOT_EXIST=不存在打印机！,
        //TITLE_EXISTED=标题的区域已经存在！,
        //TITLE_HAS_ERROR_CHARACTER=含有非法字符,区域标题只能包含字母、数字和空格，并且必须是以字母开头,
        //TITLE=标题,
        //PROTECT_SHEET_ERROR=工作表已被保护

    }
}