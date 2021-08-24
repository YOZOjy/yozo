declare namespace word {
    export interface AddIn {
        getName(): string
        delete(): void
        isCompiled(): boolean
        getPath(): string
        getIndex(): int
        setAutoload(auto: boolean): void
        isAutoload(): boolean
        setCompiled(compiled: boolean): void
        setInstalled(install: boolean): void
        isInstalled(): boolean
    }
    export interface AddIn {
        getName(): string
        delete(): void
        isCompiled(): boolean
        getPath(): string
        getIndex(): int
        setAutoload(auto: boolean): void
        isAutoload(): boolean
        setCompiled(compiled: boolean): void
        setInstalled(install: boolean): void
        isInstalled(): boolean
    }
    export interface AddIns {
        add(fileName: string, install: any): word.AddIn
        unload(removeFromList: boolean): void
        item(index: any): word.AddIn
        getCount(): int
    }
    export interface AddIns {
        add(fileName: string, install: any): word.AddIn
        unload(removeFromList: boolean): void
        item(index: any): word.AddIn
        getCount(): int
    }
    export interface Adjustments {
        getCount(): int
        getItem(index: int): double
        setItem(index: int, value: double): void
    }
    export interface Adjustments {
        getCount(): int
        getItem(index: int): double
        setItem(index: int, value: double): void
    }
    export interface AnswerWizard {
        getFiles(): word.AnswerWizardFiles
        clearFileList(): void
        resetFileList(): void
    }
    export interface AnswerWizard {
        getFiles(): word.AnswerWizardFiles
        clearFileList(): void
        resetFileList(): void
    }
    export interface AnswerWizardFiles {
        add(FileName: string): void
        delete(FileName: string): void
        item(Index: int): string
        getCount(): int
    }
    export interface AnswerWizardFiles {
        add(FileName: string): void
        delete(FileName: string): void
        item(Index: int): string
        getCount(): int
    }
    export interface Application {
        lock(eventName: string): boolean
        getAddress(name: any, addressProperties: any, useAutoText: any, displaySelectDialog: any, selectDialog: any, checkNamesDialog: any, recentAddressesChoice: any, updateRecentAddresses: any): string
        getName(): string
        getLanguage(): int
        close(saveChanges: int): void
        copy(): boolean
        getPath(): string
        getPathSeparator(): string
        resize(width: int, height: int): void
        closeAll(): int
        isValidKey(Arg1: int): boolean
        unlock(eventName: string): void
        help(HelpType: any): void
        getBuild(): string
        getWidth(): int
        getLeft(): int
        setLeft(left: int): void
        getTasks(): word.Tasks
        setTop(top: int): void
        getTop(): int
        setWidth(width: int): void
        activate(): void
        DDEPoke(Channel: number, Item: string, Data: string): void
        goBack(): void
        helpTool(): void
        Keyboard(langId: int): int
        getCommandBars(): office.CommandBars
        getDocuments(): word.Documents
        setActivePrinter(printer: string): void
        getActiveDocument(): word.Document
        getAutoCorrectEmail(): word.AutoCorrect
        getAutomationSecurity(): int
        setAutomationSecurity(security: int): void
        getBrowseExtraFileTypes(): string
        setBrowseExtraFileTypes(browseExtraFileTypes: string): void
        isArbitraryXMLSupportAvailable(): boolean
        getBackgroundPrintingStatus(): int
        getBackgroundSavingStatus(): int
        getActivePrinter(): string
        getActiveWindow(): word.Window
        getWindows(): word.Windows
        getAddIns(): word.AddIns
        getAnswerWizard(): office.AnswerWizard
        getAssistant(): office.Assistant
        getAutoCaptions(): word.AutoCaptions
        getAutoCorrect(): word.AutoCorrect
        getBrowser(): word.Browser
        isCapsLock(): boolean
        setCaption(caption: string): void
        getCaption(): string
        getCaptionLabels(): word.CaptionLabels
        isCheckLanguage(): boolean
        setCheckLanguage(checkLanguage: boolean): void
        getCOMAddIns(): office.COMAddIns
        getOptions(): word.Options
        getOptions(): office.AbstractOptions
        getDialogs(): word.Dialogs
        getDisplayAlerts(): int
        setDisplayAlerts(displayAlerts: int): void
        getEmailOptions(): word.EmailOptions
        getEmailTemplate(): string
        setEmailTemplate(emailTemplate: string): void
        getCustomDictionaries(): word.Dictionaries
        getCustomizationContext(): any
        setCustomizationContext(customizationContext: any): void
        isDefaultLegalBlackline(): boolean
        setDefaultLegalBlackline(defaultLegalBlackline: boolean): void
        getDefaultSaveFormat(): string
        setDefaultSaveFormat(defaultSaveFormat: string): void
        getDefaultTableSeparator(): string
        setDefaultTableSeparator(defaultTableSeparator: string): void
        isDisplayRecentFiles(): boolean
        setDisplayRecentFiles(displayRecentFiles: boolean): void
        isDisplayScreenTips(): boolean
        setDisplayScreenTips(displayScreenTips: boolean): void
        setDisplayDocumentInformationPanel(display: boolean): void
        isDisplayDocumentInformationPanel(): boolean
        isDisplayAutoCompleteTips(): boolean
        setDisplayAutoCompleteTips(displayAutoCompleteTips: boolean): void
        getHangulHanjaDictionaries(): word.HangulHanjaConversionDictionaries
        isDisplayScrollBars(): boolean
        setDisplayScrollBars(displayScrollBars: boolean): void
        isDisplayStatusBar(): boolean
        setDisplayStatusBar(displayStatusBar: boolean): void
        getEnableCancelKey(): int
        setEnableCancelKey(enableCancelKey: int): void
        getFeatureInstall(): int
        setFeatureInstall(featureInstall: int): void
        getFileConverters(): word.FileConverters
        isFocusInMailHeader(): boolean
        getLandscapeFontNames(): word.FontNames
        getFileDialog(fileDialogType: int): office.FileDialog
        getFileSearch(): word.FileSearch
        getFindKey(keyCode: int, keyCode2: any): word.KeyBinding
        getKeyBindings(): word.KeyBindings
        getFontNames(): word.FontNames
        getHeight(): int
        setHeight(height: int): void
        getInternational(Index: int): any
        IsObjectValid(object: any): boolean
        getKeysBoundTo(KeyCategory: int, command: string, CommandParameter: any): word.KeysBoundTo
        getLanguages(): word.Languages
        getListGalleries(): word.ListGalleries
        getMailingLabel(): word.MailingLabel
        getMailMessage(): word.MailMessage
        isMAPIAvailable(): boolean
        isMouseAvailable(): boolean
        getNewDocument(): office.NewFile
        getTemplates(): word.Templates
        isNumLock(): boolean
        isPrintPreview(): boolean
        setPrintPreview(printPreview: boolean): void
        getRecentFiles(): word.RecentFiles
        isScreenUpdating(): boolean
        isMathCoprocessorAvailable(): boolean
        getMacroContainer(): any
        getNormalTemplate(): word.Template
        getPortraitFontNames(): word.FontNames
        setScreenUpdating(screenUpdating: boolean): void
        isShowStartupDialog(): boolean
        setShowStartupDialog(showStartupDialog: boolean): void
        getSelection(): word.Selection
        getSmartTagTypes(): word.SmartTagTypes
        isSpecialMode(): boolean
        getStartupPath(): string
        setStartupPath(startupPath: string): void
        getStatusBar(): string
        setStatusBar(StatusBar: string): void
        SynonymInfo(word: string, languageID: int): word.SynonymInfo
        getSystem(): word.System
        getTaskPanes(): word.TaskPanes
        getUsableHeight(): int
        getUsableWidth(): int
        getUserAddress(): string
        setUserAddress(userAddress: string): void
        isUserControl(): boolean
        getUserInitials(): string
        setUserInitials(initials: string): void
        isVisible(): boolean
        setVisible(visible: boolean): void
        getWindowState(): int
        setWindowState(windowState: int): void
        getXMLNamespaces(): word.XMLNamespaces
        AddAddress(TagID: string, Value: string): void
        automaticChange(): void
        buildKeyCode(Arg1: int, Arg2: any, Arg3: any, Arg4: any): int
        buildKeyCode(Arg1: int, Arg2: int, Arg3: int): int
        buildKeyCode(Arg1: int): int
        buildKeyCode(Arg1: int, Arg2: int): int
        isOptionKey(Arg1: int): boolean
        isModifierKey(Arg1: int): boolean
        isShowVisualBasicEditor(): boolean
        setShowVisualBasicEditor(showVisualBasicEditor: boolean): void
        isShowWindowsInTaskbar(): boolean
        setShowWindowsInTaskbar(showWindowsInTaskbar: boolean): void
        getSmartTagRecognizers(): word.SmartTagRecognizers
        getLanguageSettings(): office.LanguageSettings
        centimetersToPoints(centimeters: double): double
        changeFileOpenDirectory(Path: string): void
        getFileOpenDirectory(): any
        checkGrammar(string: string): boolean
        checkSpelling(Word: string, customDictionary: any, ignoreUppercase: any, mainDictionary: any, CustomDictionary2: any, CustomDictionary3: any, CustomDictionary4: any, CustomDictionary5: any, CustomDictionary6: any, CustomDictionary7: any, CustomDictionary8: any, CustomDictionary9: any, CustomDictionary10: any): boolean
        cleanString(string: string): string
        DDEExecute(channel: int, command: string): void
        DDETerminate(channel: int): void
        DDEInitiate(app: string, topic: string): int
        DDERequest(Channel: number, Item: string): string
        DDETerminateAll(): void
        getDefaultTheme(DocumentType: int): string
        goForward(): void
        inchesToPoints(inches: double): double
        keyboardBidi(): void
        keyboardLatin(): void
        keyString(keyCode: int, keyCode2: any): string
        getKeyStirng(keyCode: int): any
        listCommands(listAllCommands: boolean): void
        linesToPoints(Lines: double): double
        mountVolume(zone: string, server: string, volume: string, user: any, userPassword: string, volumePassword: any): int
        newWindow(): word.Window
        nextLetter(): void
        getDefaultWebOptions(): word.DefaultWebOptions
        getSpellingSuggestions(word: string, customDictionary: any, ignoreUppercase: any, mainDictionary: any, suggestionMode: any, customDictionary2: any, customDictionary3: any, customDictionary4: any, customDictionary5: any, customDictionary6: any, customDictionary7: any, customDictionary8: any, customDictionary9: any, customDictionary10: any): word.SpellingSuggestions
        lookupNameProperties(name: string): void
        millimetersToPoints(millimeters: double): double
        move(left: int, top: int): void
        onTime(when: any, Name: string, tolerance: any): void
        printOut(background: any, append: any, range: any, outputFileName: any, from: any, to: any, item: any, copies: any, pages: any, pageType: any, printToFile: any, collate: any, fileName: any, activePrinterMacGX: any, manualDuplexPrint: any, printZoomColumn: any, printZoomRow: any, printZoomPaperWidth: any, printZoomPaperHeight: any): void
        repeat(times: any): boolean
        runMacro(macro: string): void
        Run(macro: string, args: long): int
        sendFax(): void
        showMe(): void
        quit(saveChanges: any, format: any, routeDocument: any): void
        removeApplicationListener(l: word.event.ApplicationListener): void
        pointsToCentimeters(points: double): double
        pointsToMillimeters(points: double): double
        putFocusInMailHeader(): void
        removeAllListeners(): void
        addApplicationListener(l: word.event.ApplicationListener): void
        addApplicationListener(l: any): void
        organizerCopy(Source: string, Destination: string, Name: string, object: int): void
        organizerDelete(Source: string, Name: string, object: int): void
        organizerRename(Source: string, Name: string, newName: string, object: int): void
        picasToPoints(picas: double): double
        pixelsToPoints(pixels: double, fVertical: any): double
        pointsToInches(points: double): double
        pointsToLines(points: double): double
        pointsToPicas(points: double): double
        pointsToPixels(Points: double, fVertical: any): double
        productCode(): string
        resetIgnoreAll(): void
        getNativeHandle(): int
        isRunMacroError(): boolean
        screenRefresh(): void
        setDefaultTheme(Name: string, DocumentType: int): void
        showClipboard(): void
        substituteFont(unavailableFont: string, substituteFont: string): void
        toggleKeyboard(): void
        removeListener(doc: word.Document): void
        getStartInfo(): long
        fireJSEvent(e: any): void
        fireEvent(eventType: int): void
        getWordBasic(): word.WordBasic
        setGlobalInfo(globalInfo: int): void
        setGlobalInfo(globalInfo: int, handle: int): void
        setNativeHandle(handle: int): void
        getGlobalInfo(): int
        setStartInfo(startInfo: long): void
        getMenu(hMenu: long, type: int, index: int): long
        getFind(): word.Find
        getText(): string
        moveLeft(): boolean
        moveUp(): boolean
        moveDown(): boolean
        paste(type: int): boolean
        setFont(type: string): boolean
        addRow(type: int): boolean
        cut(): boolean
        access$2(arg0: word.Application, arg1: any): void
        access$3(arg0: word.Application, arg1: string): boolean
        access$4(arg0: word.Application, arg1: string): void
        access$5(arg0: word.Application): any
        access$6(arg0: word.Application): boolean
        createBinderHandle(binder: any): void
        createPreviewPicForOle(): string
        getDefaultFilePath(): string
        showRevisionAuthor(author: string, bShow: boolean): boolean
        filterCertainAuthor(author: string, bAccept: boolean): boolean
        getAllMSBarNames(): string[]
        isMacroExist(doc: word.Document, module: string, macro: string): boolean
        enableRevision(boolValue: boolean): boolean
        setZoomRadio(zoomValue: int): void
        getZoomRadio(): double
        setDocumentId(docid: string): boolean
        getDocumentId(): string
        setDocumentType(type: string): boolean
        getDocumentType(): string
        insertText(text: string): boolean
        backspace(): boolean
        insertTable(tableName: string, columnCount: int, rowCount: int): boolean
        removeTable(tableName: string): boolean
        setRowHeight(tableName: string, num: int, height: double): boolean
        setColumnWidth(tableName: string, num: int, width: double): boolean
        setCellProtected(tableName: string, row: int, column: int, isProtected: boolean): boolean
        cursorToCell(tableName: string, row: int, column: int): boolean
        showRevision(enable: int): boolean
        acceptAllChanges(): boolean
        rejectAllChanges(): boolean
        isExsitField(fieldId: string): boolean
        setDocumentField(fieldId: string, value: string): boolean
        insertDocument(id: string, fileNameURL: string): boolean
        moveRight(): boolean
        enableRevisionAcceptCommand(enable: boolean): boolean
        enableRevisionRejectCommand(enable: boolean): boolean
        showRevisionByUser(userName: string): boolean
        insertDocumentField(id: string): boolean
        getAllDocumentField(): string
        deleteDocumentField(id: string): boolean
        showDocumentField(id: string, enable: boolean): boolean
        getDocumentFieldValue(id: string): string
        enableDocumentField(id: string, enable: boolean): boolean
        cursorToDocumentField(id: string, position: int): boolean
        deleteAllComments(): boolean
        setDocumentMultiField(ids: string, values: string, flags: int): boolean
        insertParagraph(): boolean
        renameDocField(oldName: string, newdName: string): boolean
        deleteRow(): boolean
        deleteParagraph(): boolean
        deleteTable(): boolean
        printRevision(enable: int): boolean
        compareDocuments(originalDocument: word.Document, revisedDocument: word.Document, destination: int, granularity: int, compareFormatting: boolean, compareCaseChanges: boolean, compareWhitespace: boolean, compareTables: boolean, compareHeaders: boolean, compareFootnotes: boolean, compareTextboxes: boolean, compareFields: boolean, compareComments: boolean, compareMoves: boolean, revisedAuthor: string, ignoreAllComparisonWarnings: boolean): word.Document
        access$10(arg0: word.Application): int
        access$11(arg0: word.Application, arg1: int): void
        access$12(arg0: word.Application, arg1: int): void
        access$13(arg0: word.Application, arg1: int): void
        access$14(arg0: word.Application, arg1: int): void
        access$15(arg0: word.Application): long
        access$7(arg0: word.Application): int
        access$8(arg0: word.Application): int
        access$9(arg0: word.Application): int
    }
    export interface Application {
        lock(eventName: string): boolean
        getAddress(name: any, addressProperties: any, useAutoText: any, displaySelectDialog: any, selectDialog: any, checkNamesDialog: any, recentAddressesChoice: any, updateRecentAddresses: any): string
        getName(): string
        getLanguage(): int
        close(saveChanges: int): void
        copy(): boolean
        getPath(): string
        getPathSeparator(): string
        resize(width: int, height: int): void
        closeAll(): int
        isValidKey(Arg1: int): boolean
        unlock(eventName: string): void
        help(HelpType: any): void
        getBuild(): string
        getWidth(): int
        getLeft(): int
        setLeft(left: int): void
        getTasks(): word.Tasks
        setTop(top: int): void
        getTop(): int
        setWidth(width: int): void
        activate(): void
        DDEPoke(Channel: number, Item: string, Data: string): void
        goBack(): void
        helpTool(): void
        Keyboard(langId: int): int
        getCommandBars(): office.CommandBars
        getDocuments(): word.Documents
        setActivePrinter(printer: string): void
        getActiveDocument(): word.Document
        getAutoCorrectEmail(): word.AutoCorrect
        getAutomationSecurity(): int
        setAutomationSecurity(security: int): void
        getBrowseExtraFileTypes(): string
        setBrowseExtraFileTypes(browseExtraFileTypes: string): void
        isArbitraryXMLSupportAvailable(): boolean
        getBackgroundPrintingStatus(): int
        getBackgroundSavingStatus(): int
        getActivePrinter(): string
        getActiveWindow(): word.Window
        getWindows(): word.Windows
        getAddIns(): word.AddIns
        getAnswerWizard(): office.AnswerWizard
        getAssistant(): office.Assistant
        getAutoCaptions(): word.AutoCaptions
        getAutoCorrect(): word.AutoCorrect
        getBrowser(): word.Browser
        isCapsLock(): boolean
        setCaption(caption: string): void
        getCaption(): string
        getCaptionLabels(): word.CaptionLabels
        isCheckLanguage(): boolean
        setCheckLanguage(checkLanguage: boolean): void
        getCOMAddIns(): office.COMAddIns
        getOptions(): word.Options
        getOptions(): office.AbstractOptions
        getDialogs(): word.Dialogs
        getDisplayAlerts(): int
        setDisplayAlerts(displayAlerts: int): void
        getEmailOptions(): word.EmailOptions
        getEmailTemplate(): string
        setEmailTemplate(emailTemplate: string): void
        getCustomDictionaries(): word.Dictionaries
        getCustomizationContext(): any
        setCustomizationContext(customizationContext: any): void
        isDefaultLegalBlackline(): boolean
        setDefaultLegalBlackline(defaultLegalBlackline: boolean): void
        getDefaultSaveFormat(): string
        setDefaultSaveFormat(defaultSaveFormat: string): void
        getDefaultTableSeparator(): string
        setDefaultTableSeparator(defaultTableSeparator: string): void
        isDisplayRecentFiles(): boolean
        setDisplayRecentFiles(displayRecentFiles: boolean): void
        isDisplayScreenTips(): boolean
        setDisplayScreenTips(displayScreenTips: boolean): void
        setDisplayDocumentInformationPanel(display: boolean): void
        isDisplayDocumentInformationPanel(): boolean
        isDisplayAutoCompleteTips(): boolean
        setDisplayAutoCompleteTips(displayAutoCompleteTips: boolean): void
        getHangulHanjaDictionaries(): word.HangulHanjaConversionDictionaries
        isDisplayScrollBars(): boolean
        setDisplayScrollBars(displayScrollBars: boolean): void
        isDisplayStatusBar(): boolean
        setDisplayStatusBar(displayStatusBar: boolean): void
        getEnableCancelKey(): int
        setEnableCancelKey(enableCancelKey: int): void
        getFeatureInstall(): int
        setFeatureInstall(featureInstall: int): void
        getFileConverters(): word.FileConverters
        isFocusInMailHeader(): boolean
        getLandscapeFontNames(): word.FontNames
        getFileDialog(fileDialogType: int): office.FileDialog
        getFileSearch(): word.FileSearch
        getFindKey(keyCode: int, keyCode2: any): word.KeyBinding
        getKeyBindings(): word.KeyBindings
        getFontNames(): word.FontNames
        getHeight(): int
        setHeight(height: int): void
        getInternational(Index: int): any
        IsObjectValid(object: any): boolean
        getKeysBoundTo(KeyCategory: int, command: string, CommandParameter: any): word.KeysBoundTo
        getLanguages(): word.Languages
        getListGalleries(): word.ListGalleries
        getMailingLabel(): word.MailingLabel
        getMailMessage(): word.MailMessage
        isMAPIAvailable(): boolean
        isMouseAvailable(): boolean
        getNewDocument(): office.NewFile
        getTemplates(): word.Templates
        isNumLock(): boolean
        isPrintPreview(): boolean
        setPrintPreview(printPreview: boolean): void
        getRecentFiles(): word.RecentFiles
        isScreenUpdating(): boolean
        isMathCoprocessorAvailable(): boolean
        getMacroContainer(): any
        getNormalTemplate(): word.Template
        getPortraitFontNames(): word.FontNames
        setScreenUpdating(screenUpdating: boolean): void
        isShowStartupDialog(): boolean
        setShowStartupDialog(showStartupDialog: boolean): void
        getSelection(): word.Selection
        getSmartTagTypes(): word.SmartTagTypes
        isSpecialMode(): boolean
        getStartupPath(): string
        setStartupPath(startupPath: string): void
        getStatusBar(): string
        setStatusBar(StatusBar: string): void
        SynonymInfo(word: string, languageID: int): word.SynonymInfo
        getSystem(): word.System
        getTaskPanes(): word.TaskPanes
        getUsableHeight(): int
        getUsableWidth(): int
        getUserAddress(): string
        setUserAddress(userAddress: string): void
        isUserControl(): boolean
        getUserInitials(): string
        setUserInitials(initials: string): void
        isVisible(): boolean
        setVisible(visible: boolean): void
        getWindowState(): int
        setWindowState(windowState: int): void
        getXMLNamespaces(): word.XMLNamespaces
        AddAddress(TagID: string, Value: string): void
        automaticChange(): void
        buildKeyCode(Arg1: int, Arg2: any, Arg3: any, Arg4: any): int
        buildKeyCode(Arg1: int, Arg2: int, Arg3: int): int
        buildKeyCode(Arg1: int): int
        buildKeyCode(Arg1: int, Arg2: int): int
        isOptionKey(Arg1: int): boolean
        isModifierKey(Arg1: int): boolean
        isShowVisualBasicEditor(): boolean
        setShowVisualBasicEditor(showVisualBasicEditor: boolean): void
        isShowWindowsInTaskbar(): boolean
        setShowWindowsInTaskbar(showWindowsInTaskbar: boolean): void
        getSmartTagRecognizers(): word.SmartTagRecognizers
        getLanguageSettings(): office.LanguageSettings
        centimetersToPoints(centimeters: double): double
        changeFileOpenDirectory(Path: string): void
        getFileOpenDirectory(): any
        checkGrammar(string: string): boolean
        checkSpelling(Word: string, customDictionary: any, ignoreUppercase: any, mainDictionary: any, CustomDictionary2: any, CustomDictionary3: any, CustomDictionary4: any, CustomDictionary5: any, CustomDictionary6: any, CustomDictionary7: any, CustomDictionary8: any, CustomDictionary9: any, CustomDictionary10: any): boolean
        cleanString(string: string): string
        DDEExecute(channel: int, command: string): void
        DDETerminate(channel: int): void
        DDEInitiate(app: string, topic: string): int
        DDERequest(Channel: number, Item: string): string
        DDETerminateAll(): void
        getDefaultTheme(DocumentType: int): string
        goForward(): void
        inchesToPoints(inches: double): double
        keyboardBidi(): void
        keyboardLatin(): void
        keyString(keyCode: int, keyCode2: any): string
        getKeyStirng(keyCode: int): any
        listCommands(listAllCommands: boolean): void
        linesToPoints(Lines: double): double
        mountVolume(zone: string, server: string, volume: string, user: any, userPassword: string, volumePassword: any): int
        newWindow(): word.Window
        nextLetter(): void
        getDefaultWebOptions(): word.DefaultWebOptions
        getSpellingSuggestions(word: string, customDictionary: any, ignoreUppercase: any, mainDictionary: any, suggestionMode: any, customDictionary2: any, customDictionary3: any, customDictionary4: any, customDictionary5: any, customDictionary6: any, customDictionary7: any, customDictionary8: any, customDictionary9: any, customDictionary10: any): word.SpellingSuggestions
        lookupNameProperties(name: string): void
        millimetersToPoints(millimeters: double): double
        move(left: int, top: int): void
        onTime(when: any, Name: string, tolerance: any): void
        printOut(background: any, append: any, range: any, outputFileName: any, from: any, to: any, item: any, copies: any, pages: any, pageType: any, printToFile: any, collate: any, fileName: any, activePrinterMacGX: any, manualDuplexPrint: any, printZoomColumn: any, printZoomRow: any, printZoomPaperWidth: any, printZoomPaperHeight: any): void
        repeat(times: any): boolean
        runMacro(macro: string): void
        Run(macro: string, args: long): int
        sendFax(): void
        showMe(): void
        quit(saveChanges: any, format: any, routeDocument: any): void
        removeApplicationListener(l: word.event.ApplicationListener): void
        pointsToCentimeters(points: double): double
        pointsToMillimeters(points: double): double
        putFocusInMailHeader(): void
        removeAllListeners(): void
        addApplicationListener(l: word.event.ApplicationListener): void
        addApplicationListener(l: any): void
        organizerCopy(Source: string, Destination: string, Name: string, object: int): void
        organizerDelete(Source: string, Name: string, object: int): void
        organizerRename(Source: string, Name: string, newName: string, object: int): void
        picasToPoints(picas: double): double
        pixelsToPoints(pixels: double, fVertical: any): double
        pointsToInches(points: double): double
        pointsToLines(points: double): double
        pointsToPicas(points: double): double
        pointsToPixels(Points: double, fVertical: any): double
        productCode(): string
        resetIgnoreAll(): void
        getNativeHandle(): int
        isRunMacroError(): boolean
        screenRefresh(): void
        setDefaultTheme(Name: string, DocumentType: int): void
        showClipboard(): void
        substituteFont(unavailableFont: string, substituteFont: string): void
        toggleKeyboard(): void
        removeListener(doc: word.Document): void
        getStartInfo(): long
        fireJSEvent(e: any): void
        fireEvent(eventType: int): void
        getWordBasic(): word.WordBasic
        setGlobalInfo(globalInfo: int): void
        setGlobalInfo(globalInfo: int, handle: int): void
        setNativeHandle(handle: int): void
        getGlobalInfo(): int
        setStartInfo(startInfo: long): void
        getMenu(hMenu: long, type: int, index: int): long
        getFind(): word.Find
        getText(): string
        moveLeft(): boolean
        moveUp(): boolean
        moveDown(): boolean
        paste(type: int): boolean
        setFont(type: string): boolean
        addRow(type: int): boolean
        cut(): boolean
        access$2(arg0: word.Application, arg1: any): void
        access$3(arg0: word.Application, arg1: string): boolean
        access$4(arg0: word.Application, arg1: string): void
        access$5(arg0: word.Application): any
        access$6(arg0: word.Application): boolean
        createBinderHandle(binder: any): void
        createPreviewPicForOle(): string
        getDefaultFilePath(): string
        showRevisionAuthor(author: string, bShow: boolean): boolean
        filterCertainAuthor(author: string, bAccept: boolean): boolean
        getAllMSBarNames(): string[]
        isMacroExist(doc: word.Document, module: string, macro: string): boolean
        enableRevision(boolValue: boolean): boolean
        setZoomRadio(zoomValue: int): void
        getZoomRadio(): double
        setDocumentId(docid: string): boolean
        getDocumentId(): string
        setDocumentType(type: string): boolean
        getDocumentType(): string
        insertText(text: string): boolean
        backspace(): boolean
        insertTable(tableName: string, columnCount: int, rowCount: int): boolean
        removeTable(tableName: string): boolean
        setRowHeight(tableName: string, num: int, height: double): boolean
        setColumnWidth(tableName: string, num: int, width: double): boolean
        setCellProtected(tableName: string, row: int, column: int, isProtected: boolean): boolean
        cursorToCell(tableName: string, row: int, column: int): boolean
        showRevision(enable: int): boolean
        acceptAllChanges(): boolean
        rejectAllChanges(): boolean
        isExsitField(fieldId: string): boolean
        setDocumentField(fieldId: string, value: string): boolean
        insertDocument(id: string, fileNameURL: string): boolean
        moveRight(): boolean
        enableRevisionAcceptCommand(enable: boolean): boolean
        enableRevisionRejectCommand(enable: boolean): boolean
        showRevisionByUser(userName: string): boolean
        insertDocumentField(id: string): boolean
        getAllDocumentField(): string
        deleteDocumentField(id: string): boolean
        showDocumentField(id: string, enable: boolean): boolean
        getDocumentFieldValue(id: string): string
        enableDocumentField(id: string, enable: boolean): boolean
        cursorToDocumentField(id: string, position: int): boolean
        deleteAllComments(): boolean
        setDocumentMultiField(ids: string, values: string, flags: int): boolean
        insertParagraph(): boolean
        renameDocField(oldName: string, newdName: string): boolean
        deleteRow(): boolean
        deleteParagraph(): boolean
        deleteTable(): boolean
        printRevision(enable: int): boolean
        compareDocuments(originalDocument: word.Document, revisedDocument: word.Document, destination: int, granularity: int, compareFormatting: boolean, compareCaseChanges: boolean, compareWhitespace: boolean, compareTables: boolean, compareHeaders: boolean, compareFootnotes: boolean, compareTextboxes: boolean, compareFields: boolean, compareComments: boolean, compareMoves: boolean, revisedAuthor: string, ignoreAllComparisonWarnings: boolean): word.Document
        access$10(arg0: word.Application): int
        access$11(arg0: word.Application, arg1: int): void
        access$12(arg0: word.Application, arg1: int): void
        access$13(arg0: word.Application, arg1: int): void
        access$14(arg0: word.Application, arg1: int): void
        access$15(arg0: word.Application): long
        access$7(arg0: word.Application): int
        access$8(arg0: word.Application): int
        access$9(arg0: word.Application): int
    }
    export interface AutoCaption {
        getName(): string
        getIndex(): int
        setAutoInsert(auto: boolean): void
        isAutoInsert(): boolean
        setCaptionLabel(label: any): void
        getCaptionLabel(): any
    }
    export interface AutoCaption {
        getName(): string
        getIndex(): int
        setAutoInsert(auto: boolean): void
        isAutoInsert(): boolean
        setCaptionLabel(label: any): void
        getCaptionLabel(): any
    }
    export interface AutoCaptions {
        item(index: any): word.AutoCaption
        getCount(): int
        cancelAutoInsert(): void
        removeAutoCaption(ole: string): void
        getAutoCaptionItemValue(): any
        setAutoCaptionItemValue(autoCaptionItemValue: any): void
    }
    export interface AutoCaptions {
        item(index: any): word.AutoCaption
        getCount(): int
        cancelAutoInsert(): void
        removeAutoCaption(ole: string): void
        getAutoCaptionItemValue(): any
        setAutoCaptionItemValue(autoCaptionItemValue: any): void
    }
    export interface AutoCorrect {
        getEntries(): word.AutoCorrectEntries
        setCorrectDays(correct: boolean): void
        isCorrectDays(): boolean
        setReplaceText(replace: boolean): void
        isReplaceText(): boolean
        setCorrectHangulAndAlphabet(correct: boolean): void
        isCorrectHangulAndAlphabet(): boolean
        setCorrectKeyboardSetting(correct: boolean): void
        setDisplayAutoCorrectOptions(display: boolean): void
        isDisplayAutoCorrectOptions(): boolean
        setHangulAndAlphabetAutoAdd(auto: boolean): void
        isHangulAndAlphabetAutoAdd(): boolean
        getHangulAndAlphabetExceptions(): word.HangulAndAlphabetExceptions
        setOtherCorrectionsAutoAdd(auto: boolean): void
        isOtherCorrectionsAutoAdd(): boolean
        getOtherCorrectionsExceptions(): word.OtherCorrectionsExceptions
        isReplaceTextFromSpellingChecker(): boolean
        getTwoInitialCapsExceptions(): word.TwoInitialCapsExceptions
        setCorrectCapsLock(lock: boolean): void
        isCorrectCapsLock(): boolean
        setCorrectInitialCaps(correct: boolean): void
        isCorrectInitialCaps(): boolean
        isCorrectKeyboardSetting(): boolean
        setCorrectSentenceCaps(correct: boolean): void
        isCorrectSentenceCaps(): boolean
        setCorrectTableCells(correct: boolean): void
        isCorrectTableCells(): boolean
        setFirstLetterAutoAdd(auto: boolean): void
        isFirstLetterAutoAdd(): boolean
        getFirstLetterExceptions(): word.FirstLetterExceptions
        setTwoInitialCapsAutoAdd(auto: boolean): void
        isTwoInitialCapsAutoAdd(): boolean
        setReplaceTextFromSpellingChecker(replace: boolean): void
    }
    export interface AutoCorrect {
        getEntries(): word.AutoCorrectEntries
        setCorrectDays(correct: boolean): void
        isCorrectDays(): boolean
        setReplaceText(replace: boolean): void
        isReplaceText(): boolean
        setCorrectHangulAndAlphabet(correct: boolean): void
        isCorrectHangulAndAlphabet(): boolean
        setCorrectKeyboardSetting(correct: boolean): void
        setDisplayAutoCorrectOptions(display: boolean): void
        isDisplayAutoCorrectOptions(): boolean
        setHangulAndAlphabetAutoAdd(auto: boolean): void
        isHangulAndAlphabetAutoAdd(): boolean
        getHangulAndAlphabetExceptions(): word.HangulAndAlphabetExceptions
        setOtherCorrectionsAutoAdd(auto: boolean): void
        isOtherCorrectionsAutoAdd(): boolean
        getOtherCorrectionsExceptions(): word.OtherCorrectionsExceptions
        isReplaceTextFromSpellingChecker(): boolean
        getTwoInitialCapsExceptions(): word.TwoInitialCapsExceptions
        setCorrectCapsLock(lock: boolean): void
        isCorrectCapsLock(): boolean
        setCorrectInitialCaps(correct: boolean): void
        isCorrectInitialCaps(): boolean
        isCorrectKeyboardSetting(): boolean
        setCorrectSentenceCaps(correct: boolean): void
        isCorrectSentenceCaps(): boolean
        setCorrectTableCells(correct: boolean): void
        isCorrectTableCells(): boolean
        setFirstLetterAutoAdd(auto: boolean): void
        isFirstLetterAutoAdd(): boolean
        getFirstLetterExceptions(): word.FirstLetterExceptions
        setTwoInitialCapsAutoAdd(auto: boolean): void
        isTwoInitialCapsAutoAdd(): boolean
        setReplaceTextFromSpellingChecker(replace: boolean): void
    }
    export interface AutoCorrectEntries {
        add(name: string, value: string): word.AutoCorrectEntry
        item(index: any): word.AutoCorrectEntry
        getCount(): int
        addRichText(name: string, range: word.Range): word.AutoCorrectEntry
        getAllAutoCorrectEntry(): any
        setAutoCorrectEntryItemName(name: string): void
        setAutoCorrectEntryItemValue(value: string): void
        getAutoCorrectEntryItemName(): string
        getAutoCorrectEntryItemValue(): string
    }
    export interface AutoCorrectEntries {
        add(name: string, value: string): word.AutoCorrectEntry
        item(index: any): word.AutoCorrectEntry
        getCount(): int
        addRichText(name: string, range: word.Range): word.AutoCorrectEntry
        getAllAutoCorrectEntry(): any
        setAutoCorrectEntryItemName(name: string): void
        setAutoCorrectEntryItemValue(value: string): void
        getAutoCorrectEntryItemName(): string
        getAutoCorrectEntryItemValue(): string
    }
    export interface AutoCorrectEntry {
        getName(): string
        getValue(): string
        apply(range: word.Range): void
        delete(): void
        setName(name: string): void
        setValue(value: string): void
        getIndex(): int
        isRichText(): boolean
    }
    export interface AutoCorrectEntry {
        getName(): string
        getValue(): string
        apply(range: word.Range): void
        delete(): void
        setName(name: string): void
        setValue(value: string): void
        getIndex(): int
        isRichText(): boolean
    }
    export interface AutoTextEntries {
        add(name: string, range: word.Range): word.AutoTextEntry
        item(index: any): word.AutoTextEntry
        getCount(): int
        appendToSpike(range: word.Range): word.AutoTextEntry
        setAutoTextEntryValue(autoTextEntryValue: string): void
        getAllAutoTextEntry(): any
        getAutoTextEntryValue(): string
    }
    export interface AutoTextEntries {
        add(name: string, range: word.Range): word.AutoTextEntry
        item(index: any): word.AutoTextEntry
        getCount(): int
        appendToSpike(range: word.Range): word.AutoTextEntry
        setAutoTextEntryValue(autoTextEntryValue: string): void
        getAllAutoTextEntry(): any
        getAutoTextEntryValue(): string
    }
    export interface AutoTextEntry {
        getName(): string
        getValue(): string
        delete(): void
        setName(name: string): void
        setValue(value: string): void
        getIndex(): int
        Insert(where: word.Range, richText: boolean): word.Range
        getStyleName(): string
    }
    export interface AutoTextEntry {
        getName(): string
        getValue(): string
        delete(): void
        setName(name: string): void
        setValue(value: string): void
        getIndex(): int
        Insert(where: word.Range, richText: boolean): word.Range
        getStyleName(): string
    }
    export interface Bookmark {
        getName(): string
        isEmpty(): boolean
        delete(): void
        copy(name: string): word.Bookmark
        getStart(): int
        setStart(start: int): void
        setStart(start: long): void
        setEnd(end: int): void
        setEnd(end: long): void
        isColumn(): boolean
        getEnd(): int
        getRange(): word.Range
        select(): void
        getMBookmark(): any
        getStoryType(): int
    }
    export interface Bookmark {
        getName(): string
        isEmpty(): boolean
        delete(): void
        copy(name: string): word.Bookmark
        getStart(): int
        setStart(start: int): void
        setStart(start: long): void
        setEnd(end: int): void
        setEnd(end: long): void
        isColumn(): boolean
        getEnd(): int
        getRange(): word.Range
        select(): void
        getMBookmark(): any
        getStoryType(): int
    }
    export interface Bookmarks {
        add(name: string, range: any): word.Bookmark
        remove(bookmarkName: string): void
        exists(name: string): boolean
        inRange(bk: word.Bookmark, range: word.Range): boolean
        item(index: any): word.Bookmark
        getCount(): int
        getAllBookmarks(): word.Bookmark[]
        initSpecial(): void
        getDefaultSorting(): int
        setDefaultSorting(sort: int): void
        getMBookmarks(): any
        createBookmark(mbookmark: any): word.Bookmark
        existSpecialBK(name: string): boolean
        getBookmarks(index: int): word.Bookmark
        createBookmark2(name: string): word.Bookmark
        createBookmark3(name: string): word.Bookmark
        setShowHidden(show: boolean): void
        isShowHidden(): boolean
    }
    export interface Bookmarks {
        add(name: string, range: any): word.Bookmark
        remove(bookmarkName: string): void
        exists(name: string): boolean
        inRange(bk: word.Bookmark, range: word.Range): boolean
        item(index: any): word.Bookmark
        getCount(): int
        getAllBookmarks(): word.Bookmark[]
        initSpecial(): void
        getDefaultSorting(): int
        setDefaultSorting(sort: int): void
        getMBookmarks(): any
        createBookmark(mbookmark: any): word.Bookmark
        existSpecialBK(name: string): boolean
        getBookmarks(index: int): word.Bookmark
        createBookmark2(name: string): word.Bookmark
        createBookmark3(name: string): word.Bookmark
        setShowHidden(show: boolean): void
        isShowHidden(): boolean
    }
    export interface Border {
        setColor(colorType: int): void
        getWidth(width: int): float
        isVisible(): boolean
        setVisible(visible: boolean): void
        getColor(): int
        isInside(): boolean
        setArtStyle(style: int): void
        getArtStyle(): int
        setArtWidth(width: int): void
        getArtWidth(): int
        getBordersAttr(): any
        getWidthConstants(width: float): int
        setColorIndex(index: int): void
        getColorIndex(): int
        setLineStyle(style: int): void
        getLineStyle(): int
        setLineWidth(width: int): void
        getLineWidth(): int
        setBorderVisible(visible: boolean): void
        initBorder(): void
    }
    export interface Border {
        setColor(colorType: int): void
        getWidth(width: int): float
        isVisible(): boolean
        setVisible(visible: boolean): void
        getColor(): int
        isInside(): boolean
        setArtStyle(style: int): void
        getArtStyle(): int
        setArtWidth(width: int): void
        getArtWidth(): int
        getBordersAttr(): any
        getWidthConstants(width: float): int
        setColorIndex(index: int): void
        getColorIndex(): int
        setLineStyle(style: int): void
        getLineStyle(): int
        setLineWidth(width: int): void
        getLineWidth(): int
        setBorderVisible(visible: boolean): void
        initBorder(): void
    }
    export interface Borders {
        item(type: int): word.Border
        getCount(): int
        applyPageBordersToAllSections(): void
        setEnableFirstPageInSection(enable: boolean): void
        isEnableFirstPageInSection(): boolean
        setEnableOtherPagesInSection(enable: boolean): void
        isShadow(): boolean
        setDistanceFromTop(value: int): void
        setInsideColorIndex(index: int): void
        getInsideColorIndex(): int
        setInsideLineStyle(style: int): void
        getInsideLineStyle(): int
        setInsideLineWidth(type: int): void
        getInsideLineWidth(): int
        setOutsideColorIndex(index: int): void
        getOutsideColorIndex(): int
        setOutsideLineStyle(style: any): void
        getOutsideLineStyle(): int
        setOutsideLineWidth(width: int): void
        getOutsideLineWidth(): int
        setSurroundFooter(footer: boolean): void
        setSurroundHeader(header: boolean): void
        isEnableOtherPagesInSection(): boolean
        getMBorderAttribute(): any
        getMPageBorderAttribute(): any
        getDistanceFromBottom(): int
        setDistanceFromBottom(value: int): void
        applyForPageBorder(): void
        getDistanceFromLeft(): int
        setDistanceFromLeft(value: int): void
        getDistanceFromRight(): int
        setDistanceFromRight(value: int): void
        getDistanceFromTop(): int
        applyForBorder(): void
        setAlwaysInFront(front: boolean): void
        isAlwaysInFront(): boolean
        checkPageBorder(): void
        getDistanceFrom(): int
        setDistanceFrom(type: int): void
        getEnable(): int
        setEnable(value: int): void
        isHasHorizontal(): boolean
        isHasVertical(): boolean
        setInsideColor(colorType: int): void
        getInsideColor(): int
        setJoinBorders(join: boolean): void
        isJoinBorders(): boolean
        setOutsideColor(color: int): void
        getOutsideColor(): int
        setShadow(shadow: boolean): void
        isSurroundFooter(): boolean
        isSurroundHeader(): boolean
    }
    export interface Borders {
        item(type: int): word.Border
        getCount(): int
        applyPageBordersToAllSections(): void
        setEnableFirstPageInSection(enable: boolean): void
        isEnableFirstPageInSection(): boolean
        setEnableOtherPagesInSection(enable: boolean): void
        isShadow(): boolean
        setDistanceFromTop(value: int): void
        setInsideColorIndex(index: int): void
        getInsideColorIndex(): int
        setInsideLineStyle(style: int): void
        getInsideLineStyle(): int
        setInsideLineWidth(type: int): void
        getInsideLineWidth(): int
        setOutsideColorIndex(index: int): void
        getOutsideColorIndex(): int
        setOutsideLineStyle(style: any): void
        getOutsideLineStyle(): int
        setOutsideLineWidth(width: int): void
        getOutsideLineWidth(): int
        setSurroundFooter(footer: boolean): void
        setSurroundHeader(header: boolean): void
        isEnableOtherPagesInSection(): boolean
        getMBorderAttribute(): any
        getMPageBorderAttribute(): any
        getDistanceFromBottom(): int
        setDistanceFromBottom(value: int): void
        applyForPageBorder(): void
        getDistanceFromLeft(): int
        setDistanceFromLeft(value: int): void
        getDistanceFromRight(): int
        setDistanceFromRight(value: int): void
        getDistanceFromTop(): int
        applyForBorder(): void
        setAlwaysInFront(front: boolean): void
        isAlwaysInFront(): boolean
        checkPageBorder(): void
        getDistanceFrom(): int
        setDistanceFrom(type: int): void
        getEnable(): int
        setEnable(value: int): void
        isHasHorizontal(): boolean
        isHasVertical(): boolean
        setInsideColor(colorType: int): void
        getInsideColor(): int
        setJoinBorders(join: boolean): void
        isJoinBorders(): boolean
        setOutsideColor(color: int): void
        getOutsideColor(): int
        setShadow(shadow: boolean): void
        isSurroundFooter(): boolean
        isSurroundHeader(): boolean
    }
    export interface Break {
        getRange(): word.Range
        getPageIndex(): int
    }
    export interface Break {
        getRange(): word.Range
        getPageIndex(): int
    }
    export interface Breaks {
        item(index: int): word.Break
        getCount(): int
    }
    export interface Breaks {
        item(index: int): word.Break
        getCount(): int
    }
    export interface Browser {
        next(): void
        getTarget(): int
        setTarget(type: int): void
        previous(): void
    }
    export interface Browser {
        next(): void
        getTarget(): int
        setTarget(type: int): void
        previous(): void
    }
    export interface BuildingBlock {
        getName(): string
        getValue(): string
        delete(): void
        setValue(value: string): void
        getType(): int
        insert(where: word.Range, richText: any): any
        getID(): string
        getIndex(): int
        getInsertOptions(): string
        setInsertOptions(options: string): void
        getCategory(): any
        getDescription(): string
        setDescription(value: string): void
    }
    export interface BuildingBlock {
        getName(): string
        getValue(): string
        delete(): void
        setValue(value: string): void
        getType(): int
        insert(where: word.Range, richText: any): any
        getID(): string
        getIndex(): int
        getInsertOptions(): string
        setInsertOptions(options: string): void
        getCategory(): any
        getDescription(): string
        setDescription(value: string): void
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
        getGap(): double
        presetDrop(DropType: int): void
        setAccent(accent: int): void
        setAutoAttach(auto: boolean): void
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
        getGap(): double
        presetDrop(DropType: int): void
        setAccent(accent: int): void
        setAutoAttach(auto: boolean): void
        isAutoAttach(): boolean
        isAutoLength(): boolean
        automaticLength(): void
        customDrop(drop: double): void
        customLength(length: double): void
        getDropType(): int
    }
    export interface CanvasShapes {
        range(index: any): word.ShapeRange
        item(index: any): word.Shape
        getCount(): int
        addCallout(Type: int, Left: double, Top: double, Width: double, Height: double): word.Shape
        addConnector(Type: int, BeginX: double, BeginY: double, EndX: double, EndY: double): word.Shape
        addPicture(fileName: string, linkToFile: boolean, saveWithDocument: boolean, left: double, top: double, width: double, height: double): word.Shape
        addPolyline(safeArrayOfPoints: any): word.Shape
        addTextbox(orientation: int, left: double, top: double, width: double, height: double): word.Shape
        addTextEffect(presetTextEffect: int, text: string, fontName: string, fontSize: double, fontBold: int, fontItalic: int, left: double, top: double): word.Shape
        buildFreeform(editingType: int, x1: double, y1: double): word.FreeformBuilder
        selectAll(): void
        addCurve(safeArrayOfPoints: any): word.Shape
        addLabel(Orientation: int, Left: double, Top: double, Width: double, Height: double): word.Shape
        addLine(BeginX: double, BeginY: double, EndX: double, EndY: double): word.Shape
        addShape(Type: int, left: double, top: double, width: double, height: double): word.Shape
    }
    export interface CanvasShapes {
        range(index: any): word.ShapeRange
        item(index: any): word.Shape
        getCount(): int
        addCallout(Type: int, Left: double, Top: double, Width: double, Height: double): word.Shape
        addConnector(Type: int, BeginX: double, BeginY: double, EndX: double, EndY: double): word.Shape
        addPicture(fileName: string, linkToFile: boolean, saveWithDocument: boolean, left: double, top: double, width: double, height: double): word.Shape
        addPolyline(safeArrayOfPoints: any): word.Shape
        addTextbox(orientation: int, left: double, top: double, width: double, height: double): word.Shape
        addTextEffect(presetTextEffect: int, text: string, fontName: string, fontSize: double, fontBold: int, fontItalic: int, left: double, top: double): word.Shape
        buildFreeform(editingType: int, x1: double, y1: double): word.FreeformBuilder
        selectAll(): void
        addCurve(safeArrayOfPoints: any): word.Shape
        addLabel(Orientation: int, Left: double, Top: double, Width: double, Height: double): word.Shape
        addLine(BeginX: double, BeginY: double, EndX: double, EndY: double): word.Shape
        addShape(Type: int, left: double, top: double, width: double, height: double): word.Shape
    }
    export interface CaptionLabel {
        getName(): string
        delete(): void
        getSeparator(): int
        getID(): int
        setChapterStyleLevel(level: int): void
        getChapterStyleLevel(): int
        setIncludeChapterNumber(include: boolean): void
        isIncludeChapterNumber(): boolean
        builtIn(): boolean
        setNumberStyle(style: int): void
        getNumberStyle(): int
        setPosition(position: int): void
        getPosition(): int
        setSeparator(Separator: int): void
    }
    export interface CaptionLabel {
        getName(): string
        delete(): void
        getSeparator(): int
        getID(): int
        setChapterStyleLevel(level: int): void
        getChapterStyleLevel(): int
        setIncludeChapterNumber(include: boolean): void
        isIncludeChapterNumber(): boolean
        builtIn(): boolean
        setNumberStyle(style: int): void
        getNumberStyle(): int
        setPosition(position: int): void
        getPosition(): int
        setSeparator(Separator: int): void
    }
    export interface CaptionLabels {
        add(name: string): word.CaptionLabel
        item(index: any): word.CaptionLabel
        getCount(): int
        getCaptionLabelName(): string
        setCaptionLabelName(captionLabelName: string): void
    }
    export interface CaptionLabels {
        add(name: string): word.CaptionLabel
        item(index: any): word.CaptionLabel
        getCount(): int
        getCaptionLabelName(): string
        setCaptionLabelName(captionLabelName: string): void
    }
    export interface Cell {
        split(rows: int, cols: int): void
        split(rows: any, cols: any): void
        delete(shiftCells: any): void
        delete(shiftCells: int): void
        merge(mergeTo: word.Cell): void
        getID(): string
        getWidth(): double
        setWidth(width: double): void
        setWidth(columnWidth: float, rulerStyle: int): void
        getHeight(): double
        setHeight(height: float): void
        setHeight(rowHeight: float, heightRule: int): void
        setHeight(rowHeight: any, heightRule: int): void
        getMRange(): any
        setHeightRule(rule: int): void
        getHeightRule(): int
        setLeftPadding(padding: double): void
        getLeftPadding(): double
        getNestingLevel(): int
        getTables(): word.Tables
        getPrevious(): word.Cell
        getRange(): word.Range
        select(): void
        setPreferredWidth(width: double): void
        getPreferredWidth(): double
        setPreferredWidthType(type: int): void
        getPreferredWidthType(): int
        autoSum(): void
        formula(formula: string, numFormat: string): void
        formula(formula: any, numFormat: any): void
        setID(id: string): void
        getNext(): word.Cell
        getRow(): word.Row
        getMTable(): any
        getBorders(): word.Borders
        setBottomPadding(padding: float): void
        getBottomPadding(): double
        getColumn(): word.Column
        checkIsNull(): void
        getColumnIndex(): int
        setFitText(fit: boolean): void
        isFitText(): boolean
        getRowIndex(): int
        getShading(): word.Shading
        setRightPadding(padding: double): void
        getRightPadding(): double
        setTopPadding(padding: double): void
        setVerticalAlignment(alignment: int): void
        getVerticalAlignment(): int
        getTopPadding(): double
        setWordWrap(wrap: boolean): void
        isWordWrap(): boolean
    }
    export interface Cell {
        split(rows: int, cols: int): void
        split(rows: any, cols: any): void
        delete(shiftCells: any): void
        delete(shiftCells: int): void
        merge(mergeTo: word.Cell): void
        getID(): string
        getWidth(): double
        setWidth(width: double): void
        setWidth(columnWidth: float, rulerStyle: int): void
        getHeight(): double
        setHeight(height: float): void
        setHeight(rowHeight: float, heightRule: int): void
        setHeight(rowHeight: any, heightRule: int): void
        getMRange(): any
        setHeightRule(rule: int): void
        getHeightRule(): int
        setLeftPadding(padding: double): void
        getLeftPadding(): double
        getNestingLevel(): int
        getTables(): word.Tables
        getPrevious(): word.Cell
        getRange(): word.Range
        select(): void
        setPreferredWidth(width: double): void
        getPreferredWidth(): double
        setPreferredWidthType(type: int): void
        getPreferredWidthType(): int
        autoSum(): void
        formula(formula: string, numFormat: string): void
        formula(formula: any, numFormat: any): void
        setID(id: string): void
        getNext(): word.Cell
        getRow(): word.Row
        getMTable(): any
        getBorders(): word.Borders
        setBottomPadding(padding: float): void
        getBottomPadding(): double
        getColumn(): word.Column
        checkIsNull(): void
        getColumnIndex(): int
        setFitText(fit: boolean): void
        isFitText(): boolean
        getRowIndex(): int
        getShading(): word.Shading
        setRightPadding(padding: double): void
        getRightPadding(): double
        setTopPadding(padding: double): void
        setVerticalAlignment(alignment: int): void
        getVerticalAlignment(): int
        getTopPadding(): double
        setWordWrap(wrap: boolean): void
        isWordWrap(): boolean
    }
    export interface Cells {
        add(beforeCell: word.Cell): word.Cell
        add(): word.Cell
        add(beforeCell: any): word.Cell
        split(rows: any, cols: any, mergeBeforeSplit: any): void
        split(rows: int, cols: int, mergeBeforeSplit: boolean): void
        delete(shiftCells: int): void
        delete(shiftCells: any): void
        merge(): void
        item(index: int): word.Cell
        getCount(): int
        getWidth(): double
        setWidth(width: double): void
        setWidth(columnWidth: double, rulerStyle: int): void
        getHeight(): double
        setHeight(height: double): void
        setHeight(rowHeight: double, heightRule: int): void
        setHeight(rowHeight: any, heightRule: int): void
        getMRange(): any
        setHeightRule(rule: int): void
        getHeightRule(): int
        getNestingLevel(): int
        setPreferredWidth(width: double): void
        getPreferredWidth(): float
        setPreferredWidthType(type: int): void
        getPreferredWidthType(): int
        changeCellVerticalAlignmentToEIO(Alignment: int): int
        autoFit(): void
        getMTable(): any
        getParentOfCell(): any
        distributeHeight(): void
        distributeWidth(): void
        getBorders(): word.Borders
        checkIsNull(): void
        getShading(): word.Shading
        setVerticalAlignment(alignment: int): void
        getVerticalAlignment(): int
        changeCellVerticalAlignmentToMS(Alignment: int): int
    }
    export interface Cells {
        add(beforeCell: word.Cell): word.Cell
        add(): word.Cell
        add(beforeCell: any): word.Cell
        split(rows: any, cols: any, mergeBeforeSplit: any): void
        split(rows: int, cols: int, mergeBeforeSplit: boolean): void
        delete(shiftCells: int): void
        delete(shiftCells: any): void
        merge(): void
        item(index: int): word.Cell
        getCount(): int
        getWidth(): double
        setWidth(width: double): void
        setWidth(columnWidth: double, rulerStyle: int): void
        getHeight(): double
        setHeight(height: double): void
        setHeight(rowHeight: double, heightRule: int): void
        setHeight(rowHeight: any, heightRule: int): void
        getMRange(): any
        setHeightRule(rule: int): void
        getHeightRule(): int
        getNestingLevel(): int
        setPreferredWidth(width: double): void
        getPreferredWidth(): float
        setPreferredWidthType(type: int): void
        getPreferredWidthType(): int
        changeCellVerticalAlignmentToEIO(Alignment: int): int
        autoFit(): void
        getMTable(): any
        getParentOfCell(): any
        distributeHeight(): void
        distributeWidth(): void
        getBorders(): word.Borders
        checkIsNull(): void
        getShading(): word.Shading
        setVerticalAlignment(alignment: int): void
        getVerticalAlignment(): int
        changeCellVerticalAlignmentToMS(Alignment: int): int
    }
    export interface Characters {
        getFirst(): word.Range
        item(index: any): word.Range
        getLast(): word.Range
        getCount(): int
    }
    export interface Characters {
        getFirst(): word.Range
        item(index: any): word.Range
        getLast(): word.Range
        getCount(): int
    }
    export interface CheckBox {
        getValue(): boolean
        isDefault(): boolean
        setValue(value: boolean): void
        getSize(): double
        setSize(size: double): void
        setDefault(auto: boolean): void
        isValid(): boolean
        setAutoSize(auto: boolean): void
        isAutoSize(): boolean
    }
    export interface CheckBox {
        getValue(): boolean
        isDefault(): boolean
        setValue(value: boolean): void
        getSize(): double
        setSize(size: double): void
        setDefault(auto: boolean): void
        isValid(): boolean
        setAutoSize(auto: boolean): void
        isAutoSize(): boolean
    }
    export interface ColorFormat {
        getName(): string
        setName(name: string): void
        getType(): int
        getRGB(): int
        getBlack(): int
        setBlack(black: int): void
        getCyan(): int
        setCyan(cyan: int): void
        setInk(ink: float): void
        getInk(): double
        setRGB(black: int): void
        setRGB2(black: any): void
        setCMYK(cyan: int, magenta: int, yellow: int, black: int): void
        getTintAndShade(): double
        setTintAndShade(shade: double): void
        setYellow(yellow: int): void
        getYellow(): int
        getBrightness(): double
        setBrightness(brightness: double): void
        getMagenta(): int
        setMagenta(magenta: int): void
        setOverPrint(over: boolean): void
        isOverPrint(): boolean
        getJavaRGB(): int
        getObjectThemeColor(): int
        setObjectThemeColor(cyan: int): void
    }
    export interface ColorFormat {
        getName(): string
        setName(name: string): void
        getType(): int
        getRGB(): int
        getBlack(): int
        setBlack(black: int): void
        getCyan(): int
        setCyan(cyan: int): void
        setInk(ink: float): void
        getInk(): double
        setRGB(black: int): void
        setRGB2(black: any): void
        setCMYK(cyan: int, magenta: int, yellow: int, black: int): void
        getTintAndShade(): double
        setTintAndShade(shade: double): void
        setYellow(yellow: int): void
        getYellow(): int
        getBrightness(): double
        setBrightness(brightness: double): void
        getMagenta(): int
        setMagenta(magenta: int): void
        setOverPrint(over: boolean): void
        isOverPrint(): boolean
        getJavaRGB(): int
        getObjectThemeColor(): int
        setObjectThemeColor(cyan: int): void
    }
    export interface Column {
        delete(): void
        sort(excludeHeader: boolean, sortFieldType: int, sortOrder: int, caseSensitive: boolean, bidiSort: boolean, ignoreThe: boolean, ignoreKashida: boolean, ignoreDiacritics: boolean, ignoreHe: boolean, languageID: int): void
        sort(excludeHeader: any, sortFieldType: any, sortOrder: any, caseSensitive: any, bidiSort: any, ignoreThe: any, ignoreKashida: any, ignoreDiacritics: any, ignoreHe: any, languageID: any): void
        getIndex(): int
        getWidth(): double
        setWidth(width: double): void
        setWidth(columnWidth: double, rulerStyle: int): void
        getMRange(): any
        getNestingLevel(): int
        getPrevious(): word.Column
        select(): void
        setPreferredWidth(width: double): void
        getPreferredWidth(): double
        setPreferredWidthType(type: int): void
        getPreferredWidthType(): int
        getNext(): word.Column
        getCells(): word.Cells
        autoFit(): void
        getMTable(): any
        Cells(): word.Cells
        isFirst(): boolean
        isLast(): boolean
        getBorders(): word.Borders
        getShading(): word.Shading
    }
    export interface Column {
        delete(): void
        sort(excludeHeader: boolean, sortFieldType: int, sortOrder: int, caseSensitive: boolean, bidiSort: boolean, ignoreThe: boolean, ignoreKashida: boolean, ignoreDiacritics: boolean, ignoreHe: boolean, languageID: int): void
        sort(excludeHeader: any, sortFieldType: any, sortOrder: any, caseSensitive: any, bidiSort: any, ignoreThe: any, ignoreKashida: any, ignoreDiacritics: any, ignoreHe: any, languageID: any): void
        getIndex(): int
        getWidth(): double
        setWidth(width: double): void
        setWidth(columnWidth: double, rulerStyle: int): void
        getMRange(): any
        getNestingLevel(): int
        getPrevious(): word.Column
        select(): void
        setPreferredWidth(width: double): void
        getPreferredWidth(): double
        setPreferredWidthType(type: int): void
        getPreferredWidthType(): int
        getNext(): word.Column
        getCells(): word.Cells
        autoFit(): void
        getMTable(): any
        Cells(): word.Cells
        isFirst(): boolean
        isLast(): boolean
        getBorders(): word.Borders
        getShading(): word.Shading
    }
    export interface Columns {
        add(beforeColumn: any): word.Column
        add(beforeColumn: word.Column): word.Column
        getFirst(): word.Column
        delete(): void
        item(index: int): word.Column
        getLast(): word.Column
        getCount(): int
        getWidth(): double
        setWidth(width: double): void
        setWidth(columnWidth: double, rulerStyle: int): void
        getMRange(): any
        getNestingLevel(): int
        select(): void
        setPreferredWidth(width: double): void
        getPreferredWidth(): double
        setPreferredWidthType(type: int): void
        getPreferredWidthType(): int
        autoFit(): void
        getMTable(): any
        distributeWidth(): void
        getBorders(): word.Borders
        getShading(): word.Shading
        setStartColumnIndex(startIndex: int): void
        setEndColumnIndex(endIndex: int): void
    }
    export interface Columns {
        add(beforeColumn: any): word.Column
        add(beforeColumn: word.Column): word.Column
        getFirst(): word.Column
        delete(): void
        item(index: int): word.Column
        getLast(): word.Column
        getCount(): int
        getWidth(): double
        setWidth(width: double): void
        setWidth(columnWidth: double, rulerStyle: int): void
        getMRange(): any
        getNestingLevel(): int
        select(): void
        setPreferredWidth(width: double): void
        getPreferredWidth(): double
        setPreferredWidthType(type: int): void
        getPreferredWidthType(): int
        autoFit(): void
        getMTable(): any
        distributeWidth(): void
        getBorders(): word.Borders
        getShading(): word.Shading
        setStartColumnIndex(startIndex: int): void
        setEndColumnIndex(endIndex: int): void
    }
    export interface Comment {
        delete(): void
        getDate(): any
        getIndex(): int
        getRange(): word.Range
        edit(): void
        isInk(): boolean
        getScope(): word.Range
        setAuthor(author: string): void
        getAuthor(): string
        getCommentDate(): any
        getInitial(): string
        setInitial(value: string): void
        getReference(): word.Range
        setShowTip(show: boolean): void
        isShowTip(): boolean
    }
    export interface Comment {
        delete(): void
        getDate(): any
        getIndex(): int
        getRange(): word.Range
        edit(): void
        isInk(): boolean
        getScope(): word.Range
        setAuthor(author: string): void
        getAuthor(): string
        getCommentDate(): any
        getInitial(): string
        setInitial(value: string): void
        getReference(): word.Range
        setShowTip(show: boolean): void
        isShowTip(): boolean
    }
    export interface Comments {
        add(range: word.Range, text: any): word.Comment
        item(obj: int): word.Comment
        getCount(): int
        deleteAllComments(): void
        deleteAllCommentsShown(): void
        setShowBy(name: string): void
        getShowBy(): string
    }
    export interface Comments {
        add(range: word.Range, text: any): word.Comment
        item(obj: int): word.Comment
        getCount(): int
        deleteAllComments(): void
        deleteAllCommentsShown(): void
        setShowBy(name: string): void
        getShowBy(): string
    }
    export interface ConditionalStyle {
        getFont(): word.Font
        setFont(font: word.Font): void
        setLeftPadding(padding: double): void
        getLeftPadding(): double
        getBorders(): word.Borders
        setBottomPadding(padding: double): void
        getBottomPadding(): double
        getShading(): word.Shading
        setRightPadding(padding: double): void
        getRightPadding(): double
        setTopPadding(padding: double): void
        getTopPadding(): double
        setParagraphFormat(format: word.ParagraphFormat): void
        getParagraphFormat(): word.ParagraphFormat
    }
    export interface ConditionalStyle {
        getFont(): word.Font
        setFont(font: word.Font): void
        setLeftPadding(padding: double): void
        getLeftPadding(): double
        getBorders(): word.Borders
        setBottomPadding(padding: double): void
        getBottomPadding(): double
        getShading(): word.Shading
        setRightPadding(padding: double): void
        getRightPadding(): double
        setTopPadding(padding: double): void
        getTopPadding(): double
        setParagraphFormat(format: word.ParagraphFormat): void
        getParagraphFormat(): word.ParagraphFormat
    }
    export interface ContentControl {
        delete(DeleteContents: boolean): void
        copy(): void
        getType(): int
        setColor(color: int): void
        getID(): string
        setType(type: int): void
        setTitle(title: string): void
        getTitle(): string
        cut(): void
        getColor(): int
        getRange(): word.Range
        getBuildingBlockCategory(): string
        setBuildingBlockCategory(buildingBlockCategory: string): void
        getBuildingBlockType(): int
        setBuildingBlockType(prop: int): void
        getDateCalendarType(): int
        setDateCalendarType(dateCalendarType: int): void
        getDateDisplayFormat(): string
        setDateDisplayFormat(dateDisplayFormat: string): void
        getDateDisplayLocale(): int
        setDateDisplayLocale(dateDisplayLocale: int): void
        getDateStorageFormat(): int
        setDateStorageFormat(dateStorageFormat: int): void
        getDefaultTextStyle(): any
        setDefaultTextStyle(defaultTextStyle: any): void
        getDropdownListEntries(): word.ContentControlListEntries
        isLockContentControl(): boolean
        setLockContentControl(lockContentControl: boolean): void
        getParentContentControl(): word.ContentControl
        getPlaceholderText(): word.BuildingBlock
        setPlaceholderText(buildingBlock: word.BuildingBlock, range: word.Range, text: string): void
        setUncheckedSymbol(characterNumber: char, font: string): void
        isShowingPlaceholderText(): boolean
        getAppearance(): int
        setAppearance(appearance: int): void
        isChecked(): boolean
        setChecked(checked: boolean): void
        isLockContents(): boolean
        setLockContents(lockContents: boolean): void
        isMultiLine(): boolean
        setMultiLine(multiLine: boolean): void
        setCheckedSymbol(characterNumber: int, font: string): void
        isTemporary(): boolean
        setTemporary(temporary: boolean): void
        isAllowInsertDeleteSection(): boolean
        setAllowInsertDeleteSection(allowInsertDeleteSection: boolean): void
        getRepeatingSectionItemTitle(): string
        setRepeatingSectionItemTitle(repeatingSectionItemTitle: string): void
        getLevel(): int
        getTag(): string
        setTag(tag: string): void
        ungroup(): void
    }
    export interface ContentControl {
        delete(DeleteContents: boolean): void
        copy(): void
        getType(): int
        setColor(color: int): void
        getID(): string
        setType(type: int): void
        setTitle(title: string): void
        getTitle(): string
        cut(): void
        getColor(): int
        getRange(): word.Range
        getBuildingBlockCategory(): string
        setBuildingBlockCategory(buildingBlockCategory: string): void
        getBuildingBlockType(): int
        setBuildingBlockType(prop: int): void
        getDateCalendarType(): int
        setDateCalendarType(dateCalendarType: int): void
        getDateDisplayFormat(): string
        setDateDisplayFormat(dateDisplayFormat: string): void
        getDateDisplayLocale(): int
        setDateDisplayLocale(dateDisplayLocale: int): void
        getDateStorageFormat(): int
        setDateStorageFormat(dateStorageFormat: int): void
        getDefaultTextStyle(): any
        setDefaultTextStyle(defaultTextStyle: any): void
        getDropdownListEntries(): word.ContentControlListEntries
        isLockContentControl(): boolean
        setLockContentControl(lockContentControl: boolean): void
        getParentContentControl(): word.ContentControl
        getPlaceholderText(): word.BuildingBlock
        setPlaceholderText(buildingBlock: word.BuildingBlock, range: word.Range, text: string): void
        setUncheckedSymbol(characterNumber: char, font: string): void
        isShowingPlaceholderText(): boolean
        getAppearance(): int
        setAppearance(appearance: int): void
        isChecked(): boolean
        setChecked(checked: boolean): void
        isLockContents(): boolean
        setLockContents(lockContents: boolean): void
        isMultiLine(): boolean
        setMultiLine(multiLine: boolean): void
        setCheckedSymbol(characterNumber: int, font: string): void
        isTemporary(): boolean
        setTemporary(temporary: boolean): void
        isAllowInsertDeleteSection(): boolean
        setAllowInsertDeleteSection(allowInsertDeleteSection: boolean): void
        getRepeatingSectionItemTitle(): string
        setRepeatingSectionItemTitle(repeatingSectionItemTitle: string): void
        getLevel(): int
        getTag(): string
        setTag(tag: string): void
        ungroup(): void
    }
    export interface ContentControlListEntries {
        add(text: string, value: string, index: int): word.ContentControlListEntry
        clear(): void
        item(index: int): word.ContentControlListEntry
        getCount(): int
    }
    export interface ContentControlListEntries {
        add(text: string, value: string, index: int): word.ContentControlListEntry
        clear(): void
        item(index: int): word.ContentControlListEntry
        getCount(): int
    }
    export interface ContentControlListEntry {
        getValue(): string
        delete(): void
        setValue(value: string): void
        getIndex(): int
        getText(): string
        moveUp(): void
        moveDown(): void
        setText(text: string): void
        select(): void
        setIndex(index: int): void
    }
    export interface ContentControlListEntry {
        getValue(): string
        delete(): void
        setValue(value: string): void
        getIndex(): int
        getText(): string
        moveUp(): void
        moveDown(): void
        setText(text: string): void
        select(): void
        setIndex(index: int): void
    }
    export interface ContentControls {
        add(type: int, range: any): word.ContentControl
        item(index: any): word.ContentControl
        getCount(): int
    }
    export interface ContentControls {
        add(type: int, range: any): word.ContentControl
        item(index: any): word.ContentControl
        getCount(): int
    }
    export interface CustomLabel {
        getName(): string
        delete(): void
        setName(name: string): void
        isValid(): boolean
        getIndex(): int
        getWidth(): double
        setWidth(width: double): void
        getHeight(): double
        setHeight(height: double): void
        setHorizontalPitch(pitch: double): void
        getHorizontalPitch(): double
        isDotMatrix(): boolean
        setNumberAcross(num: int): void
        getNumberAcross(): int
        setNumberDown(num: int): void
        getNumberDown(): int
        setPageSize(size: int): void
        getPageSize(): int
        setSideMargin(margin: double): void
        getSideMargin(): double
        setTopMargin(margin: double): void
        getTopMargin(): double
        setVerticalPitch(pitch: double): void
        getVerticalPitch(): double
    }
    export interface CustomLabel {
        getName(): string
        delete(): void
        setName(name: string): void
        isValid(): boolean
        getIndex(): int
        getWidth(): double
        setWidth(width: double): void
        getHeight(): double
        setHeight(height: double): void
        setHorizontalPitch(pitch: double): void
        getHorizontalPitch(): double
        isDotMatrix(): boolean
        setNumberAcross(num: int): void
        getNumberAcross(): int
        setNumberDown(num: int): void
        getNumberDown(): int
        setPageSize(size: int): void
        getPageSize(): int
        setSideMargin(margin: double): void
        getSideMargin(): double
        setTopMargin(margin: double): void
        getTopMargin(): double
        setVerticalPitch(pitch: double): void
        getVerticalPitch(): double
    }
    export interface CustomLabels {
        add(name: string, DotMatrix: any): word.CustomLabel
        item(index: any): word.CustomLabel
        getCount(): int
    }
    export interface CustomLabels {
        add(name: string, DotMatrix: any): word.CustomLabel
        item(index: any): word.CustomLabel
        getCount(): int
    }
    export interface CustomProperties {
        add(name: string, value: string): word.CustomProperty
        item(index: any): word.CustomProperty
        getCount(): int
    }
    export interface CustomProperties {
        add(name: string, value: string): word.CustomProperty
        item(index: any): word.CustomProperty
        getCount(): int
    }
    export interface CustomProperty {
        getName(): string
        getValue(): string
        delete(): void
        setValue(value: string): void
    }
    export interface CustomProperty {
        getName(): string
        getValue(): string
        delete(): void
        setValue(value: string): void
    }
    export interface DefaultWebOptions {
        getEncoding(): int
        isOptimizeForBrowser(): boolean
        setOptimizeForBrowser(optimizeForBrowser: boolean): void
        isOrganizeInFolder(): boolean
        setOrganizeInFolder(organizeInFolder: boolean): void
        setUpdateLinksOnSave(b: boolean): void
        isUpdateLinksOnSave(): boolean
        setUseLongFileNames(b: boolean): void
        isUseLongFileNames(): boolean
        getFolderSuffix(): string
        isAllowPNG(): boolean
        setAllowPNG(allowPNG: boolean): void
        getBrowserLevel(): int
        setBrowserLevel(browserLevel: int): void
        setEncoding(encoding: int): void
        setCheckIfWordIsDefaultHTMLEditor(checkIfWordIsDefaultHTMLEditor: boolean): void
        isAlwaysSaveInDefaultEncoding(): boolean
        setAlwaysSaveInDefaultEncoding(alwaysSaveInDefaultEncoding: boolean): void
        isCheckIfOfficeIsHTMLEditor(): boolean
        setCheckIfOfficeIsHTMLEditor(checkIfOfficeIsHTMLEditor: boolean): void
        isCheckIfWordIsDefaultHTMLEditor(): boolean
        isSaveNewWebPagesAsWebArchives(): boolean
        setSaveNewWebPagesAsWebArchives(saveNewWebPagesAsWebArchives: boolean): void
        getPixelsPerInch(): int
        setPixelsPerInch(pixelsPerInch: int): void
        isRelyOnCSS(): boolean
        setRelyOnCSS(relyOnCSS: boolean): void
        isRelyOnVML(): boolean
        setRelyOnVML(relyOnVML: boolean): void
        getScreenSize(): int
        setScreenSize(screenSize: int): void
        setTargetBrowser(l: int): void
        getTargetBrowser(): int
        getFonts(): office.WebPageFonts
    }
    export interface DefaultWebOptions {
        getEncoding(): int
        isOptimizeForBrowser(): boolean
        setOptimizeForBrowser(optimizeForBrowser: boolean): void
        isOrganizeInFolder(): boolean
        setOrganizeInFolder(organizeInFolder: boolean): void
        setUpdateLinksOnSave(b: boolean): void
        isUpdateLinksOnSave(): boolean
        setUseLongFileNames(b: boolean): void
        isUseLongFileNames(): boolean
        getFolderSuffix(): string
        isAllowPNG(): boolean
        setAllowPNG(allowPNG: boolean): void
        getBrowserLevel(): int
        setBrowserLevel(browserLevel: int): void
        setEncoding(encoding: int): void
        setCheckIfWordIsDefaultHTMLEditor(checkIfWordIsDefaultHTMLEditor: boolean): void
        isAlwaysSaveInDefaultEncoding(): boolean
        setAlwaysSaveInDefaultEncoding(alwaysSaveInDefaultEncoding: boolean): void
        isCheckIfOfficeIsHTMLEditor(): boolean
        setCheckIfOfficeIsHTMLEditor(checkIfOfficeIsHTMLEditor: boolean): void
        isCheckIfWordIsDefaultHTMLEditor(): boolean
        isSaveNewWebPagesAsWebArchives(): boolean
        setSaveNewWebPagesAsWebArchives(saveNewWebPagesAsWebArchives: boolean): void
        getPixelsPerInch(): int
        setPixelsPerInch(pixelsPerInch: int): void
        isRelyOnCSS(): boolean
        setRelyOnCSS(relyOnCSS: boolean): void
        isRelyOnVML(): boolean
        setRelyOnVML(relyOnVML: boolean): void
        getScreenSize(): int
        setScreenSize(screenSize: int): void
        setTargetBrowser(l: int): void
        getTargetBrowser(): int
        getFonts(): office.WebPageFonts
    }
    export interface Diagram {
        getType(): int
        convert(type: int): void
        setAutoFormat(auto: int): void
        isAutoFormat(): boolean
        setAutoLayout(auto: int): void
        isAutoLayout(): boolean
        setReverse(reverse: int): void
        isReverse(): boolean
        fitText(): void
        getNodes(): word.DiagramNodes
    }
    export interface Diagram {
        getType(): int
        convert(type: int): void
        setAutoFormat(auto: int): void
        isAutoFormat(): boolean
        setAutoLayout(auto: int): void
        isAutoLayout(): boolean
        setReverse(reverse: int): void
        isReverse(): boolean
        fitText(): void
        getNodes(): word.DiagramNodes
    }
    export interface DiagramNode {
        delete(): void
        getRoot(): word.DiagramNode
        replaceNode(targetNode: word.DiagramNode): void
        setLayout(layout: int): void
        getLayout(): int
        getShape(): word.Shape
        nextNode(): word.DiagramNode
        prevNode(): word.DiagramNode
        swapNode(targetNode: word.DiagramNode, pos: int): void
        getChildren(): word.DiagramNodeChildren
        cloneNode(copyChildren: boolean, targetNode: word.DiagramNode, pos: int): word.DiagramNode
        getDiagram(): word.Diagram
        getTextShape(): word.Shape
        transferChildren(receivingNode: word.DiagramNode): void
        addNode(pos: int, nodeType: int): word.DiagramNode
        moveNode(targetNode: word.DiagramNode, pos: int): void
    }
    export interface DiagramNode {
        delete(): void
        getRoot(): word.DiagramNode
        replaceNode(targetNode: word.DiagramNode): void
        setLayout(layout: int): void
        getLayout(): int
        getShape(): word.Shape
        nextNode(): word.DiagramNode
        prevNode(): word.DiagramNode
        swapNode(targetNode: word.DiagramNode, pos: int): void
        getChildren(): word.DiagramNodeChildren
        cloneNode(copyChildren: boolean, targetNode: word.DiagramNode, pos: int): word.DiagramNode
        getDiagram(): word.Diagram
        getTextShape(): word.Shape
        transferChildren(receivingNode: word.DiagramNode): void
        addNode(pos: int, nodeType: int): word.DiagramNode
        moveNode(targetNode: word.DiagramNode, pos: int): void
    }
    export interface DiagramNodeChildren {
        item(index: any): word.DiagramNode
        getCount(): int
        selectAll(): void
        getFirstChild(): word.DiagramNode
        getLastChild(): word.DiagramNode
        addNode(index: int, nodeType: int): word.DiagramNode
    }
    export interface DiagramNodeChildren {
        item(index: any): word.DiagramNode
        getCount(): int
        selectAll(): void
        getFirstChild(): word.DiagramNode
        getLastChild(): word.DiagramNode
        addNode(index: int, nodeType: int): word.DiagramNode
    }
    export interface DiagramNodes {
        item(index: any): word.DiagramNode
        getCount(): int
        selectAll(): void
    }
    export interface DiagramNodes {
        item(index: any): word.DiagramNode
        getCount(): int
        selectAll(): void
    }
    export interface Dialog {
        update(): void
        execute(): void
        getType(): int
        show(timeOutObj: any): int
        display(timeOutObj: any): int
        show2(timeOutObj: any): int
        show2(timeOut: int): int
        getCommandName(): string
        setDefaultTab(tab: int): void
        getDefaultTab(): int
    }
    export interface Dialog {
        update(): void
        execute(): void
        getType(): int
        show(timeOutObj: any): int
        display(timeOutObj: any): int
        show2(timeOutObj: any): int
        show2(timeOut: int): int
        getCommandName(): string
        setDefaultTab(tab: int): void
        getDefaultTab(): int
    }
    export interface Dialogs {
        item(index: int): word.Dialog
        getCount(): int
    }
    export interface Dialogs {
        item(index: int): word.Dialog
        getCount(): int
    }
    export interface Dictionaries {
        add(fileName: string): word.Dictionary
        item(index: any): word.Dictionary
        getCount(): int
        clearAll(): void
        getMaximum(): int
    }
    export interface Dictionaries {
        add(fileName: string): word.Dictionary
        item(index: any): word.Dictionary
        getCount(): int
        clearAll(): void
        getMaximum(): int
    }
    export interface Dictionary {
        getName(): string
        delete(): void
        getType(): int
        getPath(): string
        isReadOnly(): boolean
        isLanguageSpecific(): boolean
        setLanguageSpecific(languageSpecific: boolean): void
        getLanguageID(): int
        setLanguageID(languageID: int): void
    }
    export interface Dictionary {
        getName(): string
        delete(): void
        getType(): int
        getPath(): string
        isReadOnly(): boolean
        isLanguageSpecific(): boolean
        setLanguageSpecific(languageSpecific: boolean): void
        getLanguageID(): int
        setLanguageID(languageID: int): void
    }
    export interface Document {
        getName(): string
        compare(name: string, authorName: any, compareTarget: any, detectFormatChanges: any, ignoreAllComparisonWarnings: any, addToRecentFiles: any, removePersonalInformation: any, removeDateAndTime: any): void
        getFields(): word.Fields
        save(): void
        close(saveChanges: any, originalFormat: any, routeDocument: any): void
        close(saveChanges: int, originalFormat: int, routeDocument: boolean): void
        merge(fileName: string, mergeTarget: any, detectFormatChanges: any, useFormattingFrom: any, addToRecentFiles: any): void
        merge(fileName: string, mergeTarget: int, detectFormatChanges: boolean, useFormattingFrom: int, addToRecentFiles: boolean): void
        isProtected(): boolean
        getType(): int
        getPath(): string
        getPermission(): office.Permission
        getContent(): word.Range
        isReadOnly(): boolean
        setOffset(offset: long): void
        range(startObj: any, endObj: any): word.Range
        getBackground(): word.Shape
        getContainer(): any
        refresh(): void
        activate(): void
        getCommandBars(): office.CommandBars
        getActiveWindow(): word.Window
        getWindows(): word.Windows
        getMDocument(): any
        printPreview(): void
        getSelection(): word.Selection
        isUserControl(): boolean
        checkGrammar(): void
        checkSpelling(customDictionary: any, ignoreUppercase: any, alwaysSuggest: any, customDictionary2: any, customDictionary3: any, customDictionary4: any, customDictionary5: any, customDictionary6: any, customDictionary7: any, customDictionary8: any, customDictionary9: any, customDictionary10: any): void
        printOut(background: any, append: any, range: any, outputFileName: any, from: any, to: any, item: any, copies: any, pages: any, pageType: any, printToFile: any, collate: any, activePrinterMacGX: any, manualDuplexPrint: any, printZoomColumn: any, printZoomRow: any, printZoomPaperWidth: any, printZoomPaperHeight: any): void
        removeAllListeners(): void
        getNativeHandle(): int
        fireEvent(id: int): boolean
        setNativeHandle(handle: int): void
        isNew(): boolean
        saveAs(fileName: any, fileFormat: any, lockComments: any, password: any, addToRecentFiles: any, writePassword: any, readOnlyRecommended: any, embedTrueTypeFonts: any, saveNativePictureFormat: any, saveFormsData: any, saveAsAOCELetter: any, encoding: any, insertLineBreaks: any, allowSubstitutions: any, lineEnding: any, addBiDiMarks: any): void
        saveAs(fileName: string): void
        createPreviewPicForOle(): string
        getFullName(): string
        getMWorkbook(): any
        isModified(): boolean
        insertDocument(state: int, filepath: string, docName: string, withParaAttr: boolean): void
        insertDocument(state: int, fileType: int, filepath: string, docName: string, insertType: int, ignoreHeaderFooter: boolean): void
        insertDocument(state: int, filepath: string, docName: string, withParaAttr: boolean, compatSetting: boolean): void
        insertDocument(state: int, fileType: int, filepath: string, docName: string, insertType: int, ignoreHeaderFooter: boolean, compatSetting: boolean): void
        getAllDocumentField(): word.DocumentField[]
        deleteAllComments(): void
        dispose(): void
        requestFocus(): void
        getFrames(): word.Frames
        getTables(): word.Tables
        getGridDistanceVertical(): double
        setGridDistanceVertical(gridDistanceVertical: double): void
        getGridOriginHorizontal(): double
        setGridOriginHorizontal(gridOriginHorizontal: double): void
        getGridOriginVertical(): double
        setGridOriginVertical(gridOriginVertical: double): void
        getGridDistanceHorizontal(): double
        setGridDistanceHorizontal(gridDistanceHorizontal: double): void
        getRange(start: long, end: long): word.Range
        select(): void
        isSnapToGrid(): boolean
        setSnapToGrid(snapToGrid: boolean): void
        isSnapToShapes(): boolean
        setSnapToShapes(snapToShapes: boolean): void
        getPageSetup(): word.PageSetup
        initActiveCompoundEdit(): void
        fireUndoableEditUpdate(name: string): void
        getBookmarks(): word.Bookmarks
        getParagraphs(): word.Paragraphs
        getSections(): word.Sections
        deleteAllCommentsShown(): void
        acceptAllRevisions(): void
        acceptAllRevisionsShown(): void
        activeThemeDisplayName(): string
        setActiveWritingStyle(languageID: any, prop: string): void
        getActiveWritingStyle(LanguageID: any): string
        setAttachedTemplate(templateName: any): void
        getAttachedTemplate(): any
        setAutoFormatOverride(auto: boolean): void
        isAutoFormatOverride(): boolean
        setAutoHyphenation(auto: boolean): void
        isAutoHyphenation(): boolean
        checkNewSmartTags(): void
        getChildNodeSuggestions(): word.XMLChildNodeSuggestions
        closePrintPreview(): void
        computeStatistics(statistic: int, includeFootnotesAndEndnotes: any): int
        getContentControls(): word.ContentControls
        convertNumbersToText(type: int): void
        convertNumbersToText(type: any): void
        copyStylesFromTemplate(template: string): void
        countNumberedItems(numberType: any, level: any): int
        CreateLetterContent(dateFormat: string, includeHeaderFooter: boolean, pageDesign: string, letterStyle: int, letterhead: boolean, letterheadLocation: int, letterheadSize: float, recipientName: string, RecipientAddress: string, salutation: string, salutationType: int, RecipientReference: string, mailingInstructions: string, attentionLine: string, subject: string, cCList: string, returnAddress: string, senderName: string, closing: string, senderCompany: string, senderJobTitle: string, senderInitials: string, enclosureNumber: int, infoBlock: any, recipientCode: any, recipientGender: any, returnAddressShortForm: any, senderCity: any, senderCode: any, senderGender: any, senderReference: any): word.LetterContent
        getDefaultTableStyle(): any
        getDefaultTabStop(): double
        setDefaultTabStop(tabStop: double): void
        getDefaultTargetFrame(): string
        setDefaultTargetFrame(targetFrame: string): void
        deleteAllEditableRanges(editorID: any): void
        deleteAllInkAnnotations(): void
        isDisableFeatures(): boolean
        setDisableFeatures(disable: boolean): void
        setDoNotEmbedSystemFonts(embed: boolean): void
        isDoNotEmbedSystemFonts(): boolean
        setEmbedLinguisticData(embed: boolean): void
        isEmbedLinguisticData(): boolean
        setEmbedSmartTags(embed: boolean): void
        setEmbedTrueTypeFonts(embed: boolean): void
        isEmbedTrueTypeFonts(): boolean
        getFarEastLineBreakLevel(): int
        setFarEastLineBreakLevel(farEastLineBreakLevel: int): void
        setFormattingShowClear(show: boolean): void
        isFormattingShowClear(): boolean
        setFormattingShowFilter(show: int): void
        getFormattingShowFilter(): int
        setFormattingShowFont(show: boolean): void
        isFormattingShowFont(): boolean
        getCrossReferenceItems(referenceType: any): any
        setGrammarChecked(grammarChecked: boolean): void
        getGrammaticalErrors(): word.ProofreadingErrors
        isGridOriginFromMargin(): boolean
        setGridOriginFromMargin(gridOriginFromMargin: boolean): void
        setHasRoutingSlip(hasRoutingSlip: boolean): void
        getHyphenationZone(): int
        setHyphenationZone(hyphenationZone: int): void
        getJustificationMode(): int
        setJustificationMode(justificationMode: int): void
        isKerningByAlgorithm(): boolean
        setKerningByAlgorithm(kerningByAlgorithm: boolean): void
        isLanguageDetected(): boolean
        setLanguageDetected(detected: boolean): void
        makeCompatibilityDefault(): void
        manualHyphenation(): void
        setTrackRevisions(trackRevisions: boolean): void
        getNoLineBreakAfter(): string
        setNoLineBreakAfter(noLineBreakAfter: string): void
        getNoLineBreakBefore(): string
        setNoLineBreakBefore(noLineBreakBefore: string): void
        isOptimizeForWord97(): boolean
        setOptimizeForWord97(optimizeForWord97: boolean): void
        setPrintFormsData(printFormsData: boolean): void
        isPrintFractionalWidths(): boolean
        setPrintFractionalWidths(printFractionalWidths: boolean): void
        setPrintRevisions(revisions: boolean): void
        getProtectionType(): int
        getReadabilityStatistics(): word.ReadabilityStatistics
        getReadingLayoutSizeX(): int
        setReadingLayoutSizeX(readingLayoutSizeX: int): void
        getReadingLayoutSizeY(): int
        setReadingLayoutSizeY(readingLayoutSizeY: int): void
        setReadOnlyRecommended(readOnly: boolean): void
        isReadOnlyRecommended(): boolean
        rejectAllRevisions(): void
        rejectAllRevisionsShown(): void
        isRemoveDateAndTime(): boolean
        setRemoveDateAndTime(removeDateAndTime: boolean): void
        removeLockedStyles(): void
        isSaveSubsetFonts(): boolean
        setSaveSubsetFonts(saveSubsetFonts: boolean): void
        selectAllEditableRanges(editorID: any): void
        sendFaxOverInternet(recipients: any, subject: any, showMessage: any): void
        setDefaultTableStyle(Style: any, setInTemplate: boolean): void
        getInlineShapes(): word.InlineShapes
        isMasterDocument(): boolean
        isSubdocument(): boolean
        getListTemplates(): word.ListTemplates
        getMailEnvelope(): int
        getMailer(): word.Mailer
        getMailMerge(): word.MailMerge
        isTrackRevisions(): boolean
        getOpenEncoding(): int
        getPassWord(): string
        presentIt(): void
        isPrintFormsData(): boolean
        isPrintRevisions(): boolean
        setCopiesEnable(enable: boolean): void
        getRevisions(): word.Revisions
        recheckSmartTags(): void
        getProtectedArea(): word.Range[]
        removeNumbers(numberType: any): void
        removeSmartTags(): void
        removeTheme(): void
        repaginate(): void
        replyWithChanges(replay: any): void
        resetFormFields(): void
        getRoutingSlip(): word.RoutingSlip
        runAutoMacro(autoMacro: int): void
        runLetterWizard(letterContent: any, wizardMode: any): void
        setPassword(passWord: string): void
        setWritePassword(passWord: string): void
        getSaveEncoding(): int
        setSaveEncoding(saveEncoding: int): void
        isSaveFormsData(): boolean
        setSaveFormsData(saveFormsData: boolean): void
        getSaveFormat(): int
        getScripts(): office.Scripts
        selectNodes(xPath: string, prefixMapping: string, fastSearchSkippingTextNodes: boolean): word.XMLNodes
        selectSingleNode(XPath: string, prefixMapping: string, fastSearchSkippingTextNodes: boolean): word.XMLNode
        sendForReview(recipients: any, subject: any, showMessage: any, includeAttachment: any): void
        SendMailer(fileFormat: int, priority: any): void
        getSentences(): word.Sentences
        setLetterContent(letterContent: any): void
        getShapes(): word.Shapes
        isShowRevisions(): boolean
        setShowRevisions(showRevisions: boolean): void
        isShowSummary(): boolean
        setShowSummary(showSummary: boolean): void
        getSignatures(): office.SignatureSet
        getSmartDocument(): office.SmartDocument
        getSmartTags(): word.SmartTags
        getStoryRanges(): word.StoryRanges
        getStyles(): word.Styles
        getStyleSheets(): word.StyleSheets
        getSubdocuments(): word.Subdocuments
        getSummaryLength(): int
        setSummaryLength(summaryLength: int): void
        getTextEncoding(): int
        setTextEncoding(textEncoding: int): void
        undoClear(): void
        unprotect(password: any): void
        getNewPassword(password: string): string
        updateStyles(): void
        setUserControl(userControl: boolean): void
        isVBASigned(): boolean
        getVersions(): word.Versions
        getWebOptions(): word.WebOptions
        webPagePreview(): void
        isWriteReserved(): boolean
        getXMLNodes(): word.XMLNodes
        getVBProject(): vbide.VBProject
        setEnableEvents(enableEvents: boolean): void
        containMacro(name: string): boolean
        forceLayout(): void
        backspaceKey(): void
        isRemovePersonalInformation(): boolean
        setRemovePersonalInformation(removePersonalInformation: boolean): void
        setPasswordEncryptionOptions(passwordEncryptionProvider: string, passwordEncryptionAlgorithm: string, passwordEncryptionKeyLength: int, passwordEncryptionFileProperties: any): void
        getTablesOfAuthoritiesCategories(): word.TablesOfAuthoritiesCategories
        saveAs2(fileName: any, fileFormat: any, lockComments: any, password: any, addToRecentFiles: any, writePassword: any, readOnlyRecommended: any, embedTrueTypeFonts: any, saveNativePictureFormat: any, saveFormsData: any, saveAsAOCELetter: any, encoding: any, insertLineBreaks: any, allowSubstitutions: any, lineEnding: any, addBiDiMarks: any, compatibilityMode: any): void
        isSaved(): boolean
        setSaved(saved: boolean): void
        SendFax(address: string, subject: any): void
        sendMail(): void
        getSync(): office.Sync
        undo(times: any): boolean
        viewCode(): void
        getWords(): word.Words
        getSharedWorkspace(): office.SharedWorkspace
        isShowGrammaticalErrors(): boolean
        setShowGrammaticalErrors(showGrammaticalErrors: boolean): void
        isShowSpellingErrors(): boolean
        setShowSpellingErrors(showSpellingErrors: boolean): void
        isSmartTagsAsXMLProps(): boolean
        setSmartTagsAsXMLProps(smartTagsAsXMLProps: boolean): void
        isSpellingChecked(): boolean
        setSpellingChecked(spellingChecked: boolean): void
        getSpellingErrors(): word.ProofreadingErrors
        getSummaryViewMode(): int
        setSummaryViewMode(summaryViewMode: int): void
        getTablesOfAuthorities(): word.TablesOfAuthorities
        getTablesOfContents(): word.TablesOfContents
        getTablesOfFigures(): word.TablesOfFigures
        toggleFormsDesign(): void
        getTextLineEnding(): int
        setTextLineEnding(textLineEnding: int): void
        transformDocument(Path: string, dataOnly: boolean): void
        isUpdateStylesOnOpen(): boolean
        setUpdateStylesOnOpen(updateStylesOnOpen: boolean): void
        updateSummaryProperties(): void
        viewPropertyBrowser(): void
        isXMLHideNamespaces(): boolean
        setXMLHideNamespaces(hideNamespaces: boolean): void
        isXMLSaveDataOnly(): boolean
        setXMLSaveDataOnly(saveDataOnly: boolean): void
        getXMLSaveThroughXSLT(): string
        setXMLSaveThroughXSLT(saveThroughXSLT: string): void
        getXMLSchemaReferences(): word.XMLSchemaReferences
        getXMLSchemaViolations(): word.XMLNodes
        isXMLShowAdvancedErrors(): boolean
        setXMLShowAdvancedErrors(showAdvancedErrors: boolean): void
        isXMLUseXSLTWhenSaving(): boolean
        setXMLUseXSLTWhenSaving(useXSLTWhenSaving: boolean): void
        exportAsFixedFormat(outputFileName: string, exportFormat: int, openAfterExport: boolean, optimizeFor: int, range: int, from: int, to: int, item: int, includeDocProps: boolean, keepIRM: boolean, createBookmarks: int, docStructureTags: boolean, bitmapMissingFonts: boolean, useISO19005_1: boolean, fixedFormatExtClassPtr: any): void
        addDocumentListener(l: word.event.DocumentListener): void
        removeDocumentListener(l: word.event.DocumentListener): void
        getDocumentFields(): word.DocumentFields
        deleteDocumentFieldTag(name: string): void
        getPenMarkManager(): word.PenMarkManager
        getBuiltInDocumentProperties(): any
        getClickAndTypeParagraphStyle(): any
        getDisableFeaturesIntroducedAfter(): int
        setDisableFeaturesIntroducedAfter(disable: int): void
        getGridSpaceBetweenHorizontalLines(): int
        setGridSpaceBetweenHorizontalLines(gridSpaceBetweenHorizontalLines: int): void
        isPasswordEncryptionFileProperties(): boolean
        dataForm(): void
        getEmail(): word.Email
        goTo(what: any, which: any, count: any, name: any): word.Range
        goTo(what: int, which: int, count: int, name: string): word.Range
        getKind(): int
        setKind(kind: int): void
        getLists(): word.Lists
        post(): void
        protect(ranges: word.Range[], password: string, editable: boolean): void
        protect(type: int, noReset: any, password: any, useIRM: any, enforceStyleLock: any): void
        redo(times: any): boolean
        reload(): void
        reloadAs(encoding: int): void
        reply(): void
        replyAll(): void
        route(): void
        isRouted(): boolean
        isEnableEvents(): boolean
        getVariables(): word.Variables
        getActiveTheme(): string
        addToFavorites(): void
        applyTheme(name: string): void
        autoFormat(): void
        autoSummarize(length: int, mode: int, updateProperties: boolean): word.Range
        canCheckin(): boolean
        getCharacters(): word.Characters
        checkConsistency(): void
        getCodeName(): string
        getComments(): word.Comments
        setCompatibility(type: int, prop: boolean): void
        getCompatibility(type: int): boolean
        convertVietDoc(codePageOrigin: int): void
        detectLanguage(): void
        isEmbedSmartTags(): boolean
        getEndnotes(): word.Endnotes
        endReview(): void
        setEnforceStyle(enforce: boolean): void
        isEnforceStyle(): boolean
        getEnvelope(): word.Envelope
        fitToPages(): void
        followHyperlink(address: any, subAddress: any, newWindow: any, addHistory: any, extraInfo: any, method: any, headerInfo: any): void
        getFootnotes(): word.Footnotes
        getFormFields(): word.FormFields
        isFormsDesign(): boolean
        forwardMailer(): void
        getFrameset(): word.Frameset
        getLetterContent(): word.LetterContent
        isGrammarChecked(): boolean
        isHasMailer(): boolean
        setHasMailer(hasMailer: boolean): void
        isHasPassword(): boolean
        hasPassword(): boolean
        isHasRoutingSlip(): boolean
        getHTMLDivisions(): word.HTMLDivisions
        getHTMLProject(): office.HTMLProject
        getHyperlinks(): word.Hyperlinks
        isHyphenateCaps(): boolean
        setHyphenateCaps(hyphenateCaps: boolean): void
        getIndexes(): word.Indexes
        setConsecutiveHyphensLimit(limit: int): void
        getConsecutiveHyphensLimit(): int
        getCustomDocumentProperties(): any
        getDocumentLibraryVersions(): office.DocumentLibraryVersions
        getFarEastLineBreakLanguage(): int
        setFarEastLineBreakLanguage(farEastLineBreakLanguage: int): void
        setFormattingShowNumbering(show: boolean): void
        isFormattingShowNumbering(): boolean
        setFormattingShowParagraph(show: boolean): void
        isFormattingShowParagraph(): boolean
        getGridSpaceBetweenVerticalLines(): int
        setGridSpaceBetweenVerticalLines(gridSpaceBetweenVerticalLines: int): void
        getPasswordEncryptionAlgorithm(): string
        getPasswordEncryptionKeyLength(): int
        getPasswordEncryptionProvider(): string
        isPrintPostScriptOverText(): boolean
        setPrintPostScriptOverText(printPostScriptOverText: boolean): void
        isReadingModeLayoutFrozen(): boolean
        setReadingModeLayoutFrozen(readingModeLayoutFrozen: boolean): void
    }
    export interface Document {
        getName(): string
        compare(name: string, authorName: any, compareTarget: any, detectFormatChanges: any, ignoreAllComparisonWarnings: any, addToRecentFiles: any, removePersonalInformation: any, removeDateAndTime: any): void
        getFields(): word.Fields
        save(): void
        close(saveChanges: any, originalFormat: any, routeDocument: any): void
        close(saveChanges: int, originalFormat: int, routeDocument: boolean): void
        merge(fileName: string, mergeTarget: any, detectFormatChanges: any, useFormattingFrom: any, addToRecentFiles: any): void
        merge(fileName: string, mergeTarget: int, detectFormatChanges: boolean, useFormattingFrom: int, addToRecentFiles: boolean): void
        isProtected(): boolean
        getType(): int
        getPath(): string
        getPermission(): office.Permission
        getContent(): word.Range
        isReadOnly(): boolean
        setOffset(offset: long): void
        range(startObj: any, endObj: any): word.Range
        getBackground(): word.Shape
        getContainer(): any
        refresh(): void
        activate(): void
        getCommandBars(): office.CommandBars
        getActiveWindow(): word.Window
        getWindows(): word.Windows
        getMDocument(): any
        printPreview(): void
        getSelection(): word.Selection
        isUserControl(): boolean
        checkGrammar(): void
        checkSpelling(customDictionary: any, ignoreUppercase: any, alwaysSuggest: any, customDictionary2: any, customDictionary3: any, customDictionary4: any, customDictionary5: any, customDictionary6: any, customDictionary7: any, customDictionary8: any, customDictionary9: any, customDictionary10: any): void
        printOut(background: any, append: any, range: any, outputFileName: any, from: any, to: any, item: any, copies: any, pages: any, pageType: any, printToFile: any, collate: any, activePrinterMacGX: any, manualDuplexPrint: any, printZoomColumn: any, printZoomRow: any, printZoomPaperWidth: any, printZoomPaperHeight: any): void
        removeAllListeners(): void
        getNativeHandle(): int
        fireEvent(id: int): boolean
        setNativeHandle(handle: int): void
        isNew(): boolean
        saveAs(fileName: any, fileFormat: any, lockComments: any, password: any, addToRecentFiles: any, writePassword: any, readOnlyRecommended: any, embedTrueTypeFonts: any, saveNativePictureFormat: any, saveFormsData: any, saveAsAOCELetter: any, encoding: any, insertLineBreaks: any, allowSubstitutions: any, lineEnding: any, addBiDiMarks: any): void
        saveAs(fileName: string): void
        createPreviewPicForOle(): string
        getFullName(): string
        getMWorkbook(): any
        isModified(): boolean
        insertDocument(state: int, filepath: string, docName: string, withParaAttr: boolean): void
        insertDocument(state: int, fileType: int, filepath: string, docName: string, insertType: int, ignoreHeaderFooter: boolean): void
        insertDocument(state: int, filepath: string, docName: string, withParaAttr: boolean, compatSetting: boolean): void
        insertDocument(state: int, fileType: int, filepath: string, docName: string, insertType: int, ignoreHeaderFooter: boolean, compatSetting: boolean): void
        getAllDocumentField(): word.DocumentField[]
        deleteAllComments(): void
        dispose(): void
        requestFocus(): void
        getFrames(): word.Frames
        getTables(): word.Tables
        getGridDistanceVertical(): double
        setGridDistanceVertical(gridDistanceVertical: double): void
        getGridOriginHorizontal(): double
        setGridOriginHorizontal(gridOriginHorizontal: double): void
        getGridOriginVertical(): double
        setGridOriginVertical(gridOriginVertical: double): void
        getGridDistanceHorizontal(): double
        setGridDistanceHorizontal(gridDistanceHorizontal: double): void
        getRange(start: long, end: long): word.Range
        select(): void
        isSnapToGrid(): boolean
        setSnapToGrid(snapToGrid: boolean): void
        isSnapToShapes(): boolean
        setSnapToShapes(snapToShapes: boolean): void
        getPageSetup(): word.PageSetup
        initActiveCompoundEdit(): void
        fireUndoableEditUpdate(name: string): void
        getBookmarks(): word.Bookmarks
        getParagraphs(): word.Paragraphs
        getSections(): word.Sections
        deleteAllCommentsShown(): void
        acceptAllRevisions(): void
        acceptAllRevisionsShown(): void
        activeThemeDisplayName(): string
        setActiveWritingStyle(languageID: any, prop: string): void
        getActiveWritingStyle(LanguageID: any): string
        setAttachedTemplate(templateName: any): void
        getAttachedTemplate(): any
        setAutoFormatOverride(auto: boolean): void
        isAutoFormatOverride(): boolean
        setAutoHyphenation(auto: boolean): void
        isAutoHyphenation(): boolean
        checkNewSmartTags(): void
        getChildNodeSuggestions(): word.XMLChildNodeSuggestions
        closePrintPreview(): void
        computeStatistics(statistic: int, includeFootnotesAndEndnotes: any): int
        getContentControls(): word.ContentControls
        convertNumbersToText(type: int): void
        convertNumbersToText(type: any): void
        copyStylesFromTemplate(template: string): void
        countNumberedItems(numberType: any, level: any): int
        CreateLetterContent(dateFormat: string, includeHeaderFooter: boolean, pageDesign: string, letterStyle: int, letterhead: boolean, letterheadLocation: int, letterheadSize: float, recipientName: string, RecipientAddress: string, salutation: string, salutationType: int, RecipientReference: string, mailingInstructions: string, attentionLine: string, subject: string, cCList: string, returnAddress: string, senderName: string, closing: string, senderCompany: string, senderJobTitle: string, senderInitials: string, enclosureNumber: int, infoBlock: any, recipientCode: any, recipientGender: any, returnAddressShortForm: any, senderCity: any, senderCode: any, senderGender: any, senderReference: any): word.LetterContent
        getDefaultTableStyle(): any
        getDefaultTabStop(): double
        setDefaultTabStop(tabStop: double): void
        getDefaultTargetFrame(): string
        setDefaultTargetFrame(targetFrame: string): void
        deleteAllEditableRanges(editorID: any): void
        deleteAllInkAnnotations(): void
        isDisableFeatures(): boolean
        setDisableFeatures(disable: boolean): void
        setDoNotEmbedSystemFonts(embed: boolean): void
        isDoNotEmbedSystemFonts(): boolean
        setEmbedLinguisticData(embed: boolean): void
        isEmbedLinguisticData(): boolean
        setEmbedSmartTags(embed: boolean): void
        setEmbedTrueTypeFonts(embed: boolean): void
        isEmbedTrueTypeFonts(): boolean
        getFarEastLineBreakLevel(): int
        setFarEastLineBreakLevel(farEastLineBreakLevel: int): void
        setFormattingShowClear(show: boolean): void
        isFormattingShowClear(): boolean
        setFormattingShowFilter(show: int): void
        getFormattingShowFilter(): int
        setFormattingShowFont(show: boolean): void
        isFormattingShowFont(): boolean
        getCrossReferenceItems(referenceType: any): any
        setGrammarChecked(grammarChecked: boolean): void
        getGrammaticalErrors(): word.ProofreadingErrors
        isGridOriginFromMargin(): boolean
        setGridOriginFromMargin(gridOriginFromMargin: boolean): void
        setHasRoutingSlip(hasRoutingSlip: boolean): void
        getHyphenationZone(): int
        setHyphenationZone(hyphenationZone: int): void
        getJustificationMode(): int
        setJustificationMode(justificationMode: int): void
        isKerningByAlgorithm(): boolean
        setKerningByAlgorithm(kerningByAlgorithm: boolean): void
        isLanguageDetected(): boolean
        setLanguageDetected(detected: boolean): void
        makeCompatibilityDefault(): void
        manualHyphenation(): void
        setTrackRevisions(trackRevisions: boolean): void
        getNoLineBreakAfter(): string
        setNoLineBreakAfter(noLineBreakAfter: string): void
        getNoLineBreakBefore(): string
        setNoLineBreakBefore(noLineBreakBefore: string): void
        isOptimizeForWord97(): boolean
        setOptimizeForWord97(optimizeForWord97: boolean): void
        setPrintFormsData(printFormsData: boolean): void
        isPrintFractionalWidths(): boolean
        setPrintFractionalWidths(printFractionalWidths: boolean): void
        setPrintRevisions(revisions: boolean): void
        getProtectionType(): int
        getReadabilityStatistics(): word.ReadabilityStatistics
        getReadingLayoutSizeX(): int
        setReadingLayoutSizeX(readingLayoutSizeX: int): void
        getReadingLayoutSizeY(): int
        setReadingLayoutSizeY(readingLayoutSizeY: int): void
        setReadOnlyRecommended(readOnly: boolean): void
        isReadOnlyRecommended(): boolean
        rejectAllRevisions(): void
        rejectAllRevisionsShown(): void
        isRemoveDateAndTime(): boolean
        setRemoveDateAndTime(removeDateAndTime: boolean): void
        removeLockedStyles(): void
        isSaveSubsetFonts(): boolean
        setSaveSubsetFonts(saveSubsetFonts: boolean): void
        selectAllEditableRanges(editorID: any): void
        sendFaxOverInternet(recipients: any, subject: any, showMessage: any): void
        setDefaultTableStyle(Style: any, setInTemplate: boolean): void
        getInlineShapes(): word.InlineShapes
        isMasterDocument(): boolean
        isSubdocument(): boolean
        getListTemplates(): word.ListTemplates
        getMailEnvelope(): int
        getMailer(): word.Mailer
        getMailMerge(): word.MailMerge
        isTrackRevisions(): boolean
        getOpenEncoding(): int
        getPassWord(): string
        presentIt(): void
        isPrintFormsData(): boolean
        isPrintRevisions(): boolean
        setCopiesEnable(enable: boolean): void
        getRevisions(): word.Revisions
        recheckSmartTags(): void
        getProtectedArea(): word.Range[]
        removeNumbers(numberType: any): void
        removeSmartTags(): void
        removeTheme(): void
        repaginate(): void
        replyWithChanges(replay: any): void
        resetFormFields(): void
        getRoutingSlip(): word.RoutingSlip
        runAutoMacro(autoMacro: int): void
        runLetterWizard(letterContent: any, wizardMode: any): void
        setPassword(passWord: string): void
        setWritePassword(passWord: string): void
        getSaveEncoding(): int
        setSaveEncoding(saveEncoding: int): void
        isSaveFormsData(): boolean
        setSaveFormsData(saveFormsData: boolean): void
        getSaveFormat(): int
        getScripts(): office.Scripts
        selectNodes(xPath: string, prefixMapping: string, fastSearchSkippingTextNodes: boolean): word.XMLNodes
        selectSingleNode(XPath: string, prefixMapping: string, fastSearchSkippingTextNodes: boolean): word.XMLNode
        sendForReview(recipients: any, subject: any, showMessage: any, includeAttachment: any): void
        SendMailer(fileFormat: int, priority: any): void
        getSentences(): word.Sentences
        setLetterContent(letterContent: any): void
        getShapes(): word.Shapes
        isShowRevisions(): boolean
        setShowRevisions(showRevisions: boolean): void
        isShowSummary(): boolean
        setShowSummary(showSummary: boolean): void
        getSignatures(): office.SignatureSet
        getSmartDocument(): office.SmartDocument
        getSmartTags(): word.SmartTags
        getStoryRanges(): word.StoryRanges
        getStyles(): word.Styles
        getStyleSheets(): word.StyleSheets
        getSubdocuments(): word.Subdocuments
        getSummaryLength(): int
        setSummaryLength(summaryLength: int): void
        getTextEncoding(): int
        setTextEncoding(textEncoding: int): void
        undoClear(): void
        unprotect(password: any): void
        getNewPassword(password: string): string
        updateStyles(): void
        setUserControl(userControl: boolean): void
        isVBASigned(): boolean
        getVersions(): word.Versions
        getWebOptions(): word.WebOptions
        webPagePreview(): void
        isWriteReserved(): boolean
        getXMLNodes(): word.XMLNodes
        getVBProject(): vbide.VBProject
        setEnableEvents(enableEvents: boolean): void
        containMacro(name: string): boolean
        forceLayout(): void
        backspaceKey(): void
        isRemovePersonalInformation(): boolean
        setRemovePersonalInformation(removePersonalInformation: boolean): void
        setPasswordEncryptionOptions(passwordEncryptionProvider: string, passwordEncryptionAlgorithm: string, passwordEncryptionKeyLength: int, passwordEncryptionFileProperties: any): void
        getTablesOfAuthoritiesCategories(): word.TablesOfAuthoritiesCategories
        saveAs2(fileName: any, fileFormat: any, lockComments: any, password: any, addToRecentFiles: any, writePassword: any, readOnlyRecommended: any, embedTrueTypeFonts: any, saveNativePictureFormat: any, saveFormsData: any, saveAsAOCELetter: any, encoding: any, insertLineBreaks: any, allowSubstitutions: any, lineEnding: any, addBiDiMarks: any, compatibilityMode: any): void
        isSaved(): boolean
        setSaved(saved: boolean): void
        SendFax(address: string, subject: any): void
        sendMail(): void
        getSync(): office.Sync
        undo(times: any): boolean
        viewCode(): void
        getWords(): word.Words
        getSharedWorkspace(): office.SharedWorkspace
        isShowGrammaticalErrors(): boolean
        setShowGrammaticalErrors(showGrammaticalErrors: boolean): void
        isShowSpellingErrors(): boolean
        setShowSpellingErrors(showSpellingErrors: boolean): void
        isSmartTagsAsXMLProps(): boolean
        setSmartTagsAsXMLProps(smartTagsAsXMLProps: boolean): void
        isSpellingChecked(): boolean
        setSpellingChecked(spellingChecked: boolean): void
        getSpellingErrors(): word.ProofreadingErrors
        getSummaryViewMode(): int
        setSummaryViewMode(summaryViewMode: int): void
        getTablesOfAuthorities(): word.TablesOfAuthorities
        getTablesOfContents(): word.TablesOfContents
        getTablesOfFigures(): word.TablesOfFigures
        toggleFormsDesign(): void
        getTextLineEnding(): int
        setTextLineEnding(textLineEnding: int): void
        transformDocument(Path: string, dataOnly: boolean): void
        isUpdateStylesOnOpen(): boolean
        setUpdateStylesOnOpen(updateStylesOnOpen: boolean): void
        updateSummaryProperties(): void
        viewPropertyBrowser(): void
        isXMLHideNamespaces(): boolean
        setXMLHideNamespaces(hideNamespaces: boolean): void
        isXMLSaveDataOnly(): boolean
        setXMLSaveDataOnly(saveDataOnly: boolean): void
        getXMLSaveThroughXSLT(): string
        setXMLSaveThroughXSLT(saveThroughXSLT: string): void
        getXMLSchemaReferences(): word.XMLSchemaReferences
        getXMLSchemaViolations(): word.XMLNodes
        isXMLShowAdvancedErrors(): boolean
        setXMLShowAdvancedErrors(showAdvancedErrors: boolean): void
        isXMLUseXSLTWhenSaving(): boolean
        setXMLUseXSLTWhenSaving(useXSLTWhenSaving: boolean): void
        exportAsFixedFormat(outputFileName: string, exportFormat: int, openAfterExport: boolean, optimizeFor: int, range: int, from: int, to: int, item: int, includeDocProps: boolean, keepIRM: boolean, createBookmarks: int, docStructureTags: boolean, bitmapMissingFonts: boolean, useISO19005_1: boolean, fixedFormatExtClassPtr: any): void
        addDocumentListener(l: word.event.DocumentListener): void
        removeDocumentListener(l: word.event.DocumentListener): void
        getDocumentFields(): word.DocumentFields
        deleteDocumentFieldTag(name: string): void
        getPenMarkManager(): word.PenMarkManager
        getBuiltInDocumentProperties(): any
        getClickAndTypeParagraphStyle(): any
        getDisableFeaturesIntroducedAfter(): int
        setDisableFeaturesIntroducedAfter(disable: int): void
        getGridSpaceBetweenHorizontalLines(): int
        setGridSpaceBetweenHorizontalLines(gridSpaceBetweenHorizontalLines: int): void
        isPasswordEncryptionFileProperties(): boolean
        dataForm(): void
        getEmail(): word.Email
        goTo(what: any, which: any, count: any, name: any): word.Range
        goTo(what: int, which: int, count: int, name: string): word.Range
        getKind(): int
        setKind(kind: int): void
        getLists(): word.Lists
        post(): void
        protect(ranges: word.Range[], password: string, editable: boolean): void
        protect(type: int, noReset: any, password: any, useIRM: any, enforceStyleLock: any): void
        redo(times: any): boolean
        reload(): void
        reloadAs(encoding: int): void
        reply(): void
        replyAll(): void
        route(): void
        isRouted(): boolean
        isEnableEvents(): boolean
        getVariables(): word.Variables
        getActiveTheme(): string
        addToFavorites(): void
        applyTheme(name: string): void
        autoFormat(): void
        autoSummarize(length: int, mode: int, updateProperties: boolean): word.Range
        canCheckin(): boolean
        getCharacters(): word.Characters
        checkConsistency(): void
        getCodeName(): string
        getComments(): word.Comments
        setCompatibility(type: int, prop: boolean): void
        getCompatibility(type: int): boolean
        convertVietDoc(codePageOrigin: int): void
        detectLanguage(): void
        isEmbedSmartTags(): boolean
        getEndnotes(): word.Endnotes
        endReview(): void
        setEnforceStyle(enforce: boolean): void
        isEnforceStyle(): boolean
        getEnvelope(): word.Envelope
        fitToPages(): void
        followHyperlink(address: any, subAddress: any, newWindow: any, addHistory: any, extraInfo: any, method: any, headerInfo: any): void
        getFootnotes(): word.Footnotes
        getFormFields(): word.FormFields
        isFormsDesign(): boolean
        forwardMailer(): void
        getFrameset(): word.Frameset
        getLetterContent(): word.LetterContent
        isGrammarChecked(): boolean
        isHasMailer(): boolean
        setHasMailer(hasMailer: boolean): void
        isHasPassword(): boolean
        hasPassword(): boolean
        isHasRoutingSlip(): boolean
        getHTMLDivisions(): word.HTMLDivisions
        getHTMLProject(): office.HTMLProject
        getHyperlinks(): word.Hyperlinks
        isHyphenateCaps(): boolean
        setHyphenateCaps(hyphenateCaps: boolean): void
        getIndexes(): word.Indexes
        setConsecutiveHyphensLimit(limit: int): void
        getConsecutiveHyphensLimit(): int
        getCustomDocumentProperties(): any
        getDocumentLibraryVersions(): office.DocumentLibraryVersions
        getFarEastLineBreakLanguage(): int
        setFarEastLineBreakLanguage(farEastLineBreakLanguage: int): void
        setFormattingShowNumbering(show: boolean): void
        isFormattingShowNumbering(): boolean
        setFormattingShowParagraph(show: boolean): void
        isFormattingShowParagraph(): boolean
        getGridSpaceBetweenVerticalLines(): int
        setGridSpaceBetweenVerticalLines(gridSpaceBetweenVerticalLines: int): void
        getPasswordEncryptionAlgorithm(): string
        getPasswordEncryptionKeyLength(): int
        getPasswordEncryptionProvider(): string
        isPrintPostScriptOverText(): boolean
        setPrintPostScriptOverText(printPostScriptOverText: boolean): void
        isReadingModeLayoutFrozen(): boolean
        setReadingModeLayoutFrozen(readingModeLayoutFrozen: boolean): void
    }
    export interface DocumentField {
        getName(): string
        delete(): void
        setReadOnly(readOnly: boolean): void
        setName(name: string): void
        isHidden(): boolean
        isReadOnly(): boolean
        checkRange(): void
        getText(): string
        insertPicture(fileName: string): word.Shape
        setText(newValue: string): void
        getRange(): word.Range
        select(): void
        checkDocumentFieldName(name: string): void
        setHidden(isHidden: boolean): void
        isPrintOut(): boolean
        setPrintOut(printOut: boolean): void
    }
    export interface DocumentField {
        getName(): string
        delete(): void
        setReadOnly(readOnly: boolean): void
        setName(name: string): void
        isHidden(): boolean
        isReadOnly(): boolean
        checkRange(): void
        getText(): string
        insertPicture(fileName: string): word.Shape
        setText(newValue: string): void
        getRange(): word.Range
        select(): void
        checkDocumentFieldName(name: string): void
        setHidden(isHidden: boolean): void
        isPrintOut(): boolean
        setPrintOut(printOut: boolean): void
    }
    export interface DocumentFields {
        add(range: word.Range, name: string): word.DocumentField
        delete(name: string): void
        exists(name: string): boolean
        item(obj: any): word.DocumentField
        getCount(): int
        setDocumentType(documentType: string): void
        getDocumentType(): string
        setInputMode(mode: boolean): void
        getMDocumentFields(): any
        getAllDocumentFields(): word.DocumentField[]
        getDocumentField(index: int): word.DocumentField
        getDocumentField(name: string): word.DocumentField
        isInputMode(): boolean
        setDocMarker(name: string): void
        getDocMarker(): string
        setDocType(type: string): void
        getDocType(): string
    }
    export interface DocumentFields {
        add(range: word.Range, name: string): word.DocumentField
        delete(name: string): void
        exists(name: string): boolean
        item(obj: any): word.DocumentField
        getCount(): int
        setDocumentType(documentType: string): void
        getDocumentType(): string
        setInputMode(mode: boolean): void
        getMDocumentFields(): any
        getAllDocumentFields(): word.DocumentField[]
        getDocumentField(index: int): word.DocumentField
        getDocumentField(name: string): word.DocumentField
        isInputMode(): boolean
        setDocMarker(name: string): void
        getDocMarker(): string
        setDocType(type: string): void
        getDocType(): string
    }
    export interface Documents {
        add(): word.Document
        add(template: string, newTemplate: boolean, documentType: int, visible: boolean): word.Document
        add(template: any, newTemplate: any, documentType: any, visible: any): word.Document
        save(NoPrompt: any, OriginalFormat: any): void
        close(saveChanges: any, originalFormat: any, routeDocument: any): void
        open(path: string): word.Document
        open(fileName: string, confirmConversions: boolean, readOnly: boolean, addToRecentFiles: boolean, passwordDocument: string, passwordTemplate: string, revert: boolean, writePasswordDocument: string, writePasswordTemplate: string, format: int, encoding: int, visible: boolean, openAndRepair: boolean, documentDirection: int, noEncodingDialog: boolean, xMLTransform: boolean): word.Document
        open(fileNameObj: any, confirmConversionsObj: any, readOnlyObj: any, addToRecentFilesObj: any, passwordDocumentObj: any, passwordTemplateObj: any, revertObj: any, writePasswordDocumentObj: any, writePasswordTemplateObj: any, formatObj: any, encodingObj: any, visibleObj: any, openAndRepairObj: any, documentDirectionObj: any, noEncodingDialogObj: any, xMLTransformObj: any): word.Document
        item(index: any): word.Document
        getCount(): int
        getActiveDocument(): word.Document
        removeAllDocumentListeners(): void
        getDocument(mdoc: any): word.Document
        getDocument(mbook: any): word.Document
        createDocument(binder: any, mdoc: any): word.Document
        checkOut(fileName: string): void
        canCheckOut(fileName: string): boolean
        checkFile(fileName: string): string
    }
    export interface Documents {
        add(): word.Document
        add(template: string, newTemplate: boolean, documentType: int, visible: boolean): word.Document
        add(template: any, newTemplate: any, documentType: any, visible: any): word.Document
        save(NoPrompt: any, OriginalFormat: any): void
        close(saveChanges: any, originalFormat: any, routeDocument: any): void
        open(path: string): word.Document
        open(fileName: string, confirmConversions: boolean, readOnly: boolean, addToRecentFiles: boolean, passwordDocument: string, passwordTemplate: string, revert: boolean, writePasswordDocument: string, writePasswordTemplate: string, format: int, encoding: int, visible: boolean, openAndRepair: boolean, documentDirection: int, noEncodingDialog: boolean, xMLTransform: boolean): word.Document
        open(fileNameObj: any, confirmConversionsObj: any, readOnlyObj: any, addToRecentFilesObj: any, passwordDocumentObj: any, passwordTemplateObj: any, revertObj: any, writePasswordDocumentObj: any, writePasswordTemplateObj: any, formatObj: any, encodingObj: any, visibleObj: any, openAndRepairObj: any, documentDirectionObj: any, noEncodingDialogObj: any, xMLTransformObj: any): word.Document
        item(index: any): word.Document
        getCount(): int
        getActiveDocument(): word.Document
        removeAllDocumentListeners(): void
        getDocument(mdoc: any): word.Document
        getDocument(mbook: any): word.Document
        createDocument(binder: any, mdoc: any): word.Document
        checkOut(fileName: string): void
        canCheckOut(fileName: string): boolean
        checkFile(fileName: string): string
    }
    export interface DropCap {
        clear(): void
        apply(): void
        enable(): void
        setPosition(position: int): void
        getPosition(): int
        setDistanceFromText(distance: double): void
        getDistanceFromText(): double
        getFontName(): string
        setFontName(fontName: string): void
        getLinesToDrop(): int
        setLinesToDrop(lines: int): void
    }
    export interface DropCap {
        clear(): void
        apply(): void
        enable(): void
        setPosition(position: int): void
        getPosition(): int
        setDistanceFromText(distance: double): void
        getDistanceFromText(): double
        getFontName(): string
        setFontName(fontName: string): void
        getLinesToDrop(): int
        setLinesToDrop(lines: int): void
    }
    export interface DropDown {
        getValue(): int
        getDefault(): int
        setValue(v: int): void
        setDefault(v: int): void
        isValid(): boolean
        getListEntries(): word.ListEntries
    }
    export interface DropDown {
        getValue(): int
        getDefault(): int
        setValue(v: int): void
        setDefault(v: int): void
        isValid(): boolean
        getListEntries(): word.ListEntries
    }
    export interface Editor {
        getName(): string
        delete(): void
        getID(): string
        getRange(): word.Range
        selectAll(): void
        deleteAll(): void
        getNextRange(): word.Range
    }
    export interface Editor {
        getName(): string
        delete(): void
        getID(): string
        getRange(): word.Range
        selectAll(): void
        deleteAll(): void
        getNextRange(): word.Range
    }
    export interface Editors {
        add(editorID: any): word.Editor
        item(index: any): word.Editor
        getCount(): int
    }
    export interface Editors {
        add(editorID: any): word.Editor
        item(index: any): word.Editor
        getCount(): int
    }
    export interface Email {
        getCurrentEmailAuthor(): word.EmailAuthor
    }
    export interface Email {
        getCurrentEmailAuthor(): word.EmailAuthor
    }
    export interface EmailAuthor {
        getStyle(): word.Style
    }
    export interface EmailAuthor {
        getStyle(): word.Style
    }
    export interface EmailOptions {
        isAutoFormatAsYouTypeFormatListItemBeginning(): boolean
        setAutoFormatAsYouTypeFormatListItemBeginning(flag: boolean): void
        isAutoFormatAsYouTypeReplaceFarEastDashes(): boolean
        setAutoFormatAsYouTypeReplaceFarEastDashes(flag: boolean): void
        isAutoFormatAsYouTypeApplyBorders(): boolean
        setAutoFormatAsYouTypeApplyBorders(flag: boolean): void
        isAutoFormatAsYouTypeApplyBulletedLists(): boolean
        setAutoFormatAsYouTypeApplyBulletedLists(flag: boolean): void
        isAutoFormatAsYouTypeApplyClosings(): boolean
        setAutoFormatAsYouTypeApplyClosings(flag: boolean): void
        isAutoFormatAsYouTypeApplyFirstIndents(): boolean
        setAutoFormatAsYouTypeApplyFirstIndents(flag: boolean): void
        isAutoFormatAsYouTypeApplyHeadings(): boolean
        setAutoFormatAsYouTypeApplyHeadings(flag: boolean): void
        isAutoFormatAsYouTypeApplyNumberedLists(): boolean
        setAutoFormatAsYouTypeApplyNumberedLists(flag: boolean): void
        setAutoFormatAsYouTypeApplyTables(flag: boolean): void
        isAutoFormatAsYouTypeAutoLetterWizard(): boolean
        setAutoFormatAsYouTypeAutoLetterWizard(flag: boolean): void
        isAutoFormatAsYouTypeDefineStyles(): boolean
        setAutoFormatAsYouTypeDefineStyles(flag: boolean): void
        isAutoFormatAsYouTypeDeleteAutoSpaces(): boolean
        setAutoFormatAsYouTypeDeleteAutoSpaces(flag: boolean): void
        isAutoFormatAsYouTypeInsertClosings(): boolean
        setAutoFormatAsYouTypeInsertClosings(flag: boolean): void
        setAutoFormatAsYouTypeInsertOvers(flag: boolean): void
        isAutoFormatAsYouTypeApplyTables(): boolean
        isAutoFormatAsYouTypeInsertOvers(): boolean
        isAutoFormatAsYouTypeReplaceFractions(): boolean
        setAutoFormatAsYouTypeReplaceFractions(flag: boolean): void
        isAutoFormatAsYouTypeReplaceHyperlinks(): boolean
        setAutoFormatAsYouTypeReplaceHyperlinks(flag: boolean): void
        isAutoFormatAsYouTypeReplaceOrdinals(): boolean
        setAutoFormatAsYouTypeReplaceOrdinals(flag: boolean): void
        isAutoFormatAsYouTypeReplaceQuotes(): boolean
        setAutoFormatAsYouTypeReplaceQuotes(flag: boolean): void
        isAutoFormatAsYouTypeReplaceSymbols(): boolean
        setAutoFormatAsYouTypeReplaceSymbols(flag: boolean): void
        isAutoFormatAsYouTypeReplacePlainTextEmphasis(): boolean
        setAutoFormatAsYouTypeReplacePlainTextEmphasis(flag: boolean): void
        isTabIndentKey(): boolean
        setTabIndentKey(flag: boolean): void
        getEmailSignature(): word.EmailSignature
        setMarkCommentsWith(with: string): void
        getMarkCommentsWith(): string
        setNewColorOnReply(flag: boolean): void
        isNewColorOnReply(): boolean
        getPlainTextStyle(): word.Style
        setUseThemeStyleOnReply(flag: boolean): void
        isUseThemeStyleOnReply(): boolean
        getComposeStyle(): word.Style
        setEmbedSmartTag(flag: boolean): void
        isEmbedSmartTag(): boolean
        setHTMLFidelity(fidelity: int): void
        getHTMLFidelity(): int
        setMarkComments(flag: boolean): void
        isMarkComments(): boolean
        getReplyStyle(): word.Style
        setThemeName(name: string): void
        getThemeName(): string
        setUseThemeStyle(flag: boolean): void
        isUseThemeStyle(): boolean
        isRelyOnCSS(): boolean
        setRelyOnCSS(flag: boolean): void
    }
    export interface EmailOptions {
        isAutoFormatAsYouTypeFormatListItemBeginning(): boolean
        setAutoFormatAsYouTypeFormatListItemBeginning(flag: boolean): void
        isAutoFormatAsYouTypeReplaceFarEastDashes(): boolean
        setAutoFormatAsYouTypeReplaceFarEastDashes(flag: boolean): void
        isAutoFormatAsYouTypeApplyBorders(): boolean
        setAutoFormatAsYouTypeApplyBorders(flag: boolean): void
        isAutoFormatAsYouTypeApplyBulletedLists(): boolean
        setAutoFormatAsYouTypeApplyBulletedLists(flag: boolean): void
        isAutoFormatAsYouTypeApplyClosings(): boolean
        setAutoFormatAsYouTypeApplyClosings(flag: boolean): void
        isAutoFormatAsYouTypeApplyFirstIndents(): boolean
        setAutoFormatAsYouTypeApplyFirstIndents(flag: boolean): void
        isAutoFormatAsYouTypeApplyHeadings(): boolean
        setAutoFormatAsYouTypeApplyHeadings(flag: boolean): void
        isAutoFormatAsYouTypeApplyNumberedLists(): boolean
        setAutoFormatAsYouTypeApplyNumberedLists(flag: boolean): void
        setAutoFormatAsYouTypeApplyTables(flag: boolean): void
        isAutoFormatAsYouTypeAutoLetterWizard(): boolean
        setAutoFormatAsYouTypeAutoLetterWizard(flag: boolean): void
        isAutoFormatAsYouTypeDefineStyles(): boolean
        setAutoFormatAsYouTypeDefineStyles(flag: boolean): void
        isAutoFormatAsYouTypeDeleteAutoSpaces(): boolean
        setAutoFormatAsYouTypeDeleteAutoSpaces(flag: boolean): void
        isAutoFormatAsYouTypeInsertClosings(): boolean
        setAutoFormatAsYouTypeInsertClosings(flag: boolean): void
        setAutoFormatAsYouTypeInsertOvers(flag: boolean): void
        isAutoFormatAsYouTypeApplyTables(): boolean
        isAutoFormatAsYouTypeInsertOvers(): boolean
        isAutoFormatAsYouTypeReplaceFractions(): boolean
        setAutoFormatAsYouTypeReplaceFractions(flag: boolean): void
        isAutoFormatAsYouTypeReplaceHyperlinks(): boolean
        setAutoFormatAsYouTypeReplaceHyperlinks(flag: boolean): void
        isAutoFormatAsYouTypeReplaceOrdinals(): boolean
        setAutoFormatAsYouTypeReplaceOrdinals(flag: boolean): void
        isAutoFormatAsYouTypeReplaceQuotes(): boolean
        setAutoFormatAsYouTypeReplaceQuotes(flag: boolean): void
        isAutoFormatAsYouTypeReplaceSymbols(): boolean
        setAutoFormatAsYouTypeReplaceSymbols(flag: boolean): void
        isAutoFormatAsYouTypeReplacePlainTextEmphasis(): boolean
        setAutoFormatAsYouTypeReplacePlainTextEmphasis(flag: boolean): void
        isTabIndentKey(): boolean
        setTabIndentKey(flag: boolean): void
        getEmailSignature(): word.EmailSignature
        setMarkCommentsWith(with: string): void
        getMarkCommentsWith(): string
        setNewColorOnReply(flag: boolean): void
        isNewColorOnReply(): boolean
        getPlainTextStyle(): word.Style
        setUseThemeStyleOnReply(flag: boolean): void
        isUseThemeStyleOnReply(): boolean
        getComposeStyle(): word.Style
        setEmbedSmartTag(flag: boolean): void
        isEmbedSmartTag(): boolean
        setHTMLFidelity(fidelity: int): void
        getHTMLFidelity(): int
        setMarkComments(flag: boolean): void
        isMarkComments(): boolean
        getReplyStyle(): word.Style
        setThemeName(name: string): void
        getThemeName(): string
        setUseThemeStyle(flag: boolean): void
        isUseThemeStyle(): boolean
        isRelyOnCSS(): boolean
        setRelyOnCSS(flag: boolean): void
    }
    export interface EmailSignature {
        getEmailSignatureEntries(): word.EmailSignatureEntries
        setNewMessageSignature(s: string): void
        getNewMessageSignature(): string
        setReplyMessageSignature(s: string): void
        getReplyMessageSignature(): string
    }
    export interface EmailSignature {
        getEmailSignatureEntries(): word.EmailSignatureEntries
        setNewMessageSignature(s: string): void
        getNewMessageSignature(): string
        setReplyMessageSignature(s: string): void
        getReplyMessageSignature(): string
    }
    export interface EmailSignatureEntries {
        add(name: string, range: word.Range): word.EmailSignatureEntry
        item(index: int): int
        getCount(): int
    }
    export interface EmailSignatureEntries {
        add(name: string, range: word.Range): word.EmailSignatureEntry
        item(index: int): int
        getCount(): int
    }
    export interface EmailSignatureEntry {
        getName(): string
        delete(): void
        setName(s: string): void
        getIndex(): int
    }
    export interface EmailSignatureEntry {
        getName(): string
        delete(): void
        setName(s: string): void
        getIndex(): int
    }
    export interface Endnote {
        delete(): void
        getIndex(): int
        getRange(): word.Range
        getReference(): word.Range
    }
    export interface Endnote {
        delete(): void
        getIndex(): int
        getRange(): word.Range
        getReference(): word.Range
    }
    export interface EndnoteOptions {
        getLocation(): int
        setLocation(location: int): void
        setNumberStyle(numberStyle: int): void
        getNumberStyle(): int
        setStartingNumber(l: int): void
        getStartingNumber(): int
        setNumberingRule(rule: int): void
        getNumberingRule(): int
    }
    export interface EndnoteOptions {
        getLocation(): int
        setLocation(location: int): void
        setNumberStyle(numberStyle: int): void
        getNumberStyle(): int
        setStartingNumber(l: int): void
        getStartingNumber(): int
        setNumberingRule(rule: int): void
        getNumberingRule(): int
    }
    export interface Endnotes {
        add(range: word.Range, reference: any, text: any): word.Endnote
        getLocation(): int
        getSeparator(): word.Range
        convert(): void
        item(index: int): word.Endnote
        item(arg0: int): word.Note
        getCount(): int
        setLocation(location: int): void
        setNumberStyle(num: int): void
        getNumberStyle(): int
        resetContinuationSeparator(): void
        getContinuationNotice(): word.Range
        setStartingNumber(num: int): void
        getStartingNumber(): int
        getContinuationSeparator(): word.Range
        resetContinuationNotice(): void
        swapWithFootnotes(): void
        setNumberingRule(num: int): void
        getNumberingRule(): int
        resetSeparator(): void
    }
    export interface Endnotes {
        add(range: word.Range, reference: any, text: any): word.Endnote
        getLocation(): int
        getSeparator(): word.Range
        convert(): void
        item(index: int): word.Endnote
        item(arg0: int): word.Note
        getCount(): int
        setLocation(location: int): void
        setNumberStyle(num: int): void
        getNumberStyle(): int
        resetContinuationSeparator(): void
        getContinuationNotice(): word.Range
        setStartingNumber(num: int): void
        getStartingNumber(): int
        getContinuationSeparator(): word.Range
        resetContinuationNotice(): void
        swapWithFootnotes(): void
        setNumberingRule(num: int): void
        getNumberingRule(): int
        resetSeparator(): void
    }
    export interface Envelope {
        getAddress(): word.Range
        insert(extractAddress: any, address: any, autoText: any, omitReturnAddress: any, returnAddress: any, returnAutoText: any, printBarCode: any, printFIMA: any, size: any, height: any, width: any, feedSource: any, addressFromLeft: any, addressFromTop: any, returnAddressFromLeft: any, returnAddressFromTop: any, defaultFaceUp: any, defaultOrientation: any, printEPostage: any, vertical: any, recipientNamefromLeft: any, recipientNamefromTop: any, recipientPostalfromLeft: any, recipientPostalfromTop: any, senderNamefromLeft: any, senderNamefromTop: any, senderPostalfromLeft: any, senderPostalfromTop: any): void
        options(): void
        printOut(extractAddress: any, address: any, autoText: any, omitReturnAddress: any, returnAddress: any, returnAutoText: any, printBarCode: any, printFIMA: any, size: any, height: any, width: any, feedSource: any, addressFromLeft: any, addressFromTop: any, returnAddressFromLeft: any, returnAddressFromTop: any, defaultFaceUp: any, defaultOrientation: any, printEPostage: any, vertical: any, recipientNamefromLeft: any, recipientNamefromTop: any, recipientPostalfromLeft: any, recipientPostalfromTop: any, senderNamefromLeft: any, senderNamefromTop: any, senderPostalfromLeft: any, senderPostalfromTop: any): void
        setDefaultOmitReturnAddress(flag: boolean): void
        isDefaultOmitReturnAddress(): boolean
        setRecipientPostalfromLeft(f: double): void
        getRecipientPostalfromLeft(): double
        setRecipientPostalfromTop(f: double): void
        getRecipientPostalfromTop(): double
        setRecipientNamefromLeft(f: double): void
        getRecipientNamefromLeft(): double
        setRecipientNamefromTop(f: double): void
        getRecipientNamefromTop(): double
        setReturnAddressFromLeft(f: double): void
        getReturnAddressFromLeft(): double
        setReturnAddressFromTop(f: double): void
        getReturnAddressFromTop(): double
        setSenderNamefromLeft(f: double): void
        getSenderNamefromLeft(): double
        getReturnAddressStyle(): word.Style
        setSenderNamefromTop(f: double): void
        getSenderNamefromTop(): double
        setSenderPostalfromLeft(f: double): void
        getSenderPostalfromLeft(): double
        setSenderPostalfromTop(f: double): void
        getSenderPostalfromTop(): double
        setAddressFromLeft(f: double): void
        getAddressFromLeft(): double
        setAddressFromTop(f: double): void
        getAddressFromTop(): double
        setDefaultOrientation(orientation: int): void
        getDefaultOrientation(): int
        setDefaultPrintBarCode(flag: boolean): void
        isDefaultPrintBarCode(): boolean
        setDefaultPrintFIMA(flag: boolean): void
        isDefaultPrintFIMA(): boolean
        getAddressStyle(): word.Style
        setDefaultFaceUp(flag: boolean): void
        isDefaultFaceUp(): boolean
        setDefaultHeight(f: double): void
        getDefaultHeight(): double
        setDefaultWidth(f: double): void
        getDefaultWidth(): double
        setDefaultSize(s: string): void
        getDefaultSize(): string
        setFeedSource(source: int): void
        getFeedSource(): int
        updateDocument(): void
        getReturnAddress(): word.Range
        setVertical(b: boolean): void
        isVertical(): boolean
    }
    export interface Envelope {
        getAddress(): word.Range
        insert(extractAddress: any, address: any, autoText: any, omitReturnAddress: any, returnAddress: any, returnAutoText: any, printBarCode: any, printFIMA: any, size: any, height: any, width: any, feedSource: any, addressFromLeft: any, addressFromTop: any, returnAddressFromLeft: any, returnAddressFromTop: any, defaultFaceUp: any, defaultOrientation: any, printEPostage: any, vertical: any, recipientNamefromLeft: any, recipientNamefromTop: any, recipientPostalfromLeft: any, recipientPostalfromTop: any, senderNamefromLeft: any, senderNamefromTop: any, senderPostalfromLeft: any, senderPostalfromTop: any): void
        options(): void
        printOut(extractAddress: any, address: any, autoText: any, omitReturnAddress: any, returnAddress: any, returnAutoText: any, printBarCode: any, printFIMA: any, size: any, height: any, width: any, feedSource: any, addressFromLeft: any, addressFromTop: any, returnAddressFromLeft: any, returnAddressFromTop: any, defaultFaceUp: any, defaultOrientation: any, printEPostage: any, vertical: any, recipientNamefromLeft: any, recipientNamefromTop: any, recipientPostalfromLeft: any, recipientPostalfromTop: any, senderNamefromLeft: any, senderNamefromTop: any, senderPostalfromLeft: any, senderPostalfromTop: any): void
        setDefaultOmitReturnAddress(flag: boolean): void
        isDefaultOmitReturnAddress(): boolean
        setRecipientPostalfromLeft(f: double): void
        getRecipientPostalfromLeft(): double
        setRecipientPostalfromTop(f: double): void
        getRecipientPostalfromTop(): double
        setRecipientNamefromLeft(f: double): void
        getRecipientNamefromLeft(): double
        setRecipientNamefromTop(f: double): void
        getRecipientNamefromTop(): double
        setReturnAddressFromLeft(f: double): void
        getReturnAddressFromLeft(): double
        setReturnAddressFromTop(f: double): void
        getReturnAddressFromTop(): double
        setSenderNamefromLeft(f: double): void
        getSenderNamefromLeft(): double
        getReturnAddressStyle(): word.Style
        setSenderNamefromTop(f: double): void
        getSenderNamefromTop(): double
        setSenderPostalfromLeft(f: double): void
        getSenderPostalfromLeft(): double
        setSenderPostalfromTop(f: double): void
        getSenderPostalfromTop(): double
        setAddressFromLeft(f: double): void
        getAddressFromLeft(): double
        setAddressFromTop(f: double): void
        getAddressFromTop(): double
        setDefaultOrientation(orientation: int): void
        getDefaultOrientation(): int
        setDefaultPrintBarCode(flag: boolean): void
        isDefaultPrintBarCode(): boolean
        setDefaultPrintFIMA(flag: boolean): void
        isDefaultPrintFIMA(): boolean
        getAddressStyle(): word.Style
        setDefaultFaceUp(flag: boolean): void
        isDefaultFaceUp(): boolean
        setDefaultHeight(f: double): void
        getDefaultHeight(): double
        setDefaultWidth(f: double): void
        getDefaultWidth(): double
        setDefaultSize(s: string): void
        getDefaultSize(): string
        setFeedSource(source: int): void
        getFeedSource(): int
        updateDocument(): void
        getReturnAddress(): word.Range
        setVertical(b: boolean): void
        isVertical(): boolean
    }
    export interface Field {
        update(): boolean
        delete(): void
        copy(): void
        getType(): int
        getResult(): word.Range
        isLocked(): boolean
        unlink(): void
        getData(): string
        getIndex(): int
        cut(): void
        getPrevious(): word.Field
        select(): void
        getNext(): word.Field
        getOLEFormat(): word.OLEFormat
        setShowCodes(b: boolean): void
        getShowCodes(): boolean
        updateSource(): void
        getCode(): word.Range
        setCode(range: word.Range): void
        setData(s: string): void
        doClick(): void
        getMfield(): any
        getInlineShape(): word.InlineShape
        getLinkFormat(): word.LinkFormat
        setLocked(b: boolean): void
        getKind(): word.Field
    }
    export interface Field {
        update(): boolean
        delete(): void
        copy(): void
        getType(): int
        getResult(): word.Range
        isLocked(): boolean
        unlink(): void
        getData(): string
        getIndex(): int
        cut(): void
        getPrevious(): word.Field
        select(): void
        getNext(): word.Field
        getOLEFormat(): word.OLEFormat
        setShowCodes(b: boolean): void
        getShowCodes(): boolean
        updateSource(): void
        getCode(): word.Range
        setCode(range: word.Range): void
        setData(s: string): void
        doClick(): void
        getMfield(): any
        getInlineShape(): word.InlineShape
        getLinkFormat(): word.LinkFormat
        setLocked(b: boolean): void
        getKind(): word.Field
    }
    export interface Fields {
        add(range: word.Range, type: int, text: string, preserveFormatting: boolean): word.Field
        add(range: word.Range, type: any, text: any, preserveFormatting: any): word.Field
        update(): int
        getType(): int
        unlink(): void
        item(index: int): word.Field
        getCount(): int
        updateSource(): void
        getLocked(): int
        toggleShowCodes(): void
        getMFields(): any
        setLocked(v: int): void
    }
    export interface Fields {
        add(range: word.Range, type: int, text: string, preserveFormatting: boolean): word.Field
        add(range: word.Range, type: any, text: any, preserveFormatting: any): word.Field
        update(): int
        getType(): int
        unlink(): void
        item(index: int): word.Field
        getCount(): int
        updateSource(): void
        getLocked(): int
        toggleShowCodes(): void
        getMFields(): any
        setLocked(v: int): void
    }
    export interface FileConverter {
        getName(): string
        getPath(): string
        getClassName(): string
        getExtensions(): string
        isCanOpen(): boolean
        isCanSave(): boolean
        getFormatName(): string
        getOpenFormat(): int
        getSaveFormat(): int
    }
    export interface FileConverter {
        getName(): string
        getPath(): string
        getClassName(): string
        getExtensions(): string
        isCanOpen(): boolean
        isCanSave(): boolean
        getFormatName(): string
        getOpenFormat(): int
        getSaveFormat(): int
    }
    export interface FileConverters {
        item(index: int): word.FileConverter
        item(index: string): word.FileConverter
        getCount(): int
        setConvertMacWordChevrons(index: int): void
        getConvertMacWordChevrons(): int
    }
    export interface FileConverters {
        item(index: int): word.FileConverter
        item(index: string): word.FileConverter
        getCount(): int
        setConvertMacWordChevrons(index: int): void
        getConvertMacWordChevrons(): int
    }
    export interface FileSearch {
        setLastModified(mod: int): void
        getFileName(): string
        getLastModified(): int
        Execute(sortBy: int, sortOrder: int, alwaysAccurate: boolean): int
        setFileName(fileName: string): void
        setFileType(type: int): void
        getFileTypes(): office.FileTypes
        getFoundFiles(): office.FoundFiles
        getLookIn(): string
        setLookIn(s: string): void
        newSearch(): void
        getPropertyTests(): office.PropertyTests
        refreshScopes(): void
        getSearchFolders(): office.SearchFolders
        isMatchAllWordForms(): boolean
        setMatchAllWordForms(b: boolean): void
        isMatchTextExactly(): boolean
        setMatchTextExactly(b: boolean): void
        isSearchSubFolders(): boolean
        setSearchSubFolders(b: boolean): void
        getTextOrProperty(): string
        setTextOrProperty(s: string): void
        getFileType(): int
    }
    export interface FileSearch {
        setLastModified(mod: int): void
        getFileName(): string
        getLastModified(): int
        Execute(sortBy: int, sortOrder: int, alwaysAccurate: boolean): int
        setFileName(fileName: string): void
        setFileType(type: int): void
        getFileTypes(): office.FileTypes
        getFoundFiles(): office.FoundFiles
        getLookIn(): string
        setLookIn(s: string): void
        newSearch(): void
        getPropertyTests(): office.PropertyTests
        refreshScopes(): void
        getSearchFolders(): office.SearchFolders
        isMatchAllWordForms(): boolean
        setMatchAllWordForms(b: boolean): void
        isMatchTextExactly(): boolean
        setMatchTextExactly(b: boolean): void
        isSearchSubFolders(): boolean
        setSearchSubFolders(b: boolean): void
        getTextOrProperty(): string
        setTextOrProperty(s: string): void
        getFileType(): int
    }
    export interface FileTypes {
        add(fileType: int): void
        remove(index: int): void
        item(index: int): int
        getCount(): int
    }
    export interface FileTypes {
        add(fileType: int): void
        remove(index: int): void
        item(index: int): int
        getCount(): int
    }
    export interface FillFormat {
        apply(mFill: any): void
        getType(): int
        setVisible(visible: int): void
        setGradientAngle(value: double): void
        getGradientAngle(): double
        solid(): void
        getFillAttribute(): any
        getBackColor(mFill: any): word.ColorFormat
        getBackColor(): word.ColorFormat
        getForeColor(): word.ColorFormat
        getForeColor(mFill: any): word.ColorFormat
        applyBackground(mFill: any): void
        getGradientStyle(): int
        setGradientStyle(grtype: int): void
        getPattern(): int
        patterned(pattern: int): void
        presetGradient(style: int, variant: int, presetGradientType: int): void
        getPresetTexture(): int
        presetTextured(texture: int): void
        getTextureName(): string
        setTextureName(fileName: string): void
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
        SetGradientColorType(colorType: int): void
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
        setTextureOffsetX(value: double): void
        getTextureOffsetY(): double
        setTextureOffsetY(value: double): void
        getTextureVerticalScale(): double
        setTextureVerticalScale(value: double): void
        getPictureEffects(): office.PictureEffects
        getGradientVariant(): int
        getAngleFromMsStyle(msStyle: int): int
        getTextureHorizontalScale(): double
        setTextureHorizontalScale(value: double): void
    }
    export interface FillFormat {
        apply(mFill: any): void
        getType(): int
        setVisible(visible: int): void
        setGradientAngle(value: double): void
        getGradientAngle(): double
        solid(): void
        getFillAttribute(): any
        getBackColor(mFill: any): word.ColorFormat
        getBackColor(): word.ColorFormat
        getForeColor(): word.ColorFormat
        getForeColor(mFill: any): word.ColorFormat
        applyBackground(mFill: any): void
        getGradientStyle(): int
        setGradientStyle(grtype: int): void
        getPattern(): int
        patterned(pattern: int): void
        presetGradient(style: int, variant: int, presetGradientType: int): void
        getPresetTexture(): int
        presetTextured(texture: int): void
        getTextureName(): string
        setTextureName(fileName: string): void
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
        SetGradientColorType(colorType: int): void
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
        setTextureOffsetX(value: double): void
        getTextureOffsetY(): double
        setTextureOffsetY(value: double): void
        getTextureVerticalScale(): double
        setTextureVerticalScale(value: double): void
        getPictureEffects(): office.PictureEffects
        getGradientVariant(): int
        getAngleFromMsStyle(msStyle: int): int
        getTextureHorizontalScale(): double
        setTextureHorizontalScale(value: double): void
    }
    export interface Find {
        execute(findText: string, replaceWith: string): boolean
        execute(FindText: any, MatchCase: any, MatchWholeWord: any, MatchWildcards: any, MatchSoundsLike: any, MatchAllWordForms: any, Forward: any, Wrap: any, Format: any, ReplaceWith: any, Replace: any, MatchKashida: any, MatchDiacritics: any, MatchAlefHamza: any, MatchControl: any): boolean
        isLetterOrDigit(ch: char): boolean
        getFont(): word.Font
        getText(): string
        setFont(font: word.Font): void
        setText(text: string): void
        setParagraphFormat(format: word.ParagraphFormat): void
        getParagraphFormat(): word.ParagraphFormat
        getWrap(): int
        setWrap(wrap: int): void
        isFormat(): boolean
        isFound(): boolean
        isFrame(): word.Frame
        getHighlight(): int
        clearFormatting(): void
        setMatchCase(matchCase: boolean): void
        isMatchCase(): boolean
        setWholeWord(whole: boolean): void
        isWholeWord(): boolean
        getReplacement(): word.Replacement
        setMatchByte(b: boolean): void
        resetRange(range: word.Range): void
        setFormat(b: boolean): void
        setForward(b: boolean): void
        isForward(): boolean
        setHighlight(h: int): void
        isMatchAlefHamza(): boolean
        isMatchByte(): boolean
        isMatchControl(): boolean
        setMatchControl(b: boolean): void
        isMatchFuzzy(): boolean
        setMatchFuzzy(b: boolean): void
        isMatchKashida(): boolean
        setMatchKashida(b: boolean): void
        isMatchWholeWord(): boolean
        isMatchWildcards(): boolean
        getNoProofing(): int
        setNoProofing(b: int): void
        isMatchAllWordForms(): boolean
        setMatchAllWordForms(b: boolean): void
        setMatchWholeWord(b: boolean): void
        isAllLetterOrDigit(str: string): boolean
        clearAllFuzzyOptions(): void
        setCorrectHangulEndings(b: boolean): void
        isCorrectHangulEndings(): boolean
        getLanguageIDFarEast(): int
        setLanguageIDFarEast(id: int): void
        getLanguageIDOther(): int
        setLanguageIDOther(id: int): void
        setMatchAlefHamza(b: boolean): void
        isMatchDiacritics(): boolean
        setMatchDiacritics(b: boolean): void
        isMatchSoundsLike(): boolean
        setMatchSoundsLike(b: boolean): void
        setMatchWildcards(b: boolean): void
        setAllFuzzyOptions(): void
        getStyle(): any
        setStyle(style: any): void
        getLanguageID(): int
        setLanguageID(id: int): void
    }
    export interface Find {
        execute(findText: string, replaceWith: string): boolean
        execute(FindText: any, MatchCase: any, MatchWholeWord: any, MatchWildcards: any, MatchSoundsLike: any, MatchAllWordForms: any, Forward: any, Wrap: any, Format: any, ReplaceWith: any, Replace: any, MatchKashida: any, MatchDiacritics: any, MatchAlefHamza: any, MatchControl: any): boolean
        isLetterOrDigit(ch: char): boolean
        getFont(): word.Font
        getText(): string
        setFont(font: word.Font): void
        setText(text: string): void
        setParagraphFormat(format: word.ParagraphFormat): void
        getParagraphFormat(): word.ParagraphFormat
        getWrap(): int
        setWrap(wrap: int): void
        isFormat(): boolean
        isFound(): boolean
        isFrame(): word.Frame
        getHighlight(): int
        clearFormatting(): void
        setMatchCase(matchCase: boolean): void
        isMatchCase(): boolean
        setWholeWord(whole: boolean): void
        isWholeWord(): boolean
        getReplacement(): word.Replacement
        setMatchByte(b: boolean): void
        resetRange(range: word.Range): void
        setFormat(b: boolean): void
        setForward(b: boolean): void
        isForward(): boolean
        setHighlight(h: int): void
        isMatchAlefHamza(): boolean
        isMatchByte(): boolean
        isMatchControl(): boolean
        setMatchControl(b: boolean): void
        isMatchFuzzy(): boolean
        setMatchFuzzy(b: boolean): void
        isMatchKashida(): boolean
        setMatchKashida(b: boolean): void
        isMatchWholeWord(): boolean
        isMatchWildcards(): boolean
        getNoProofing(): int
        setNoProofing(b: int): void
        isMatchAllWordForms(): boolean
        setMatchAllWordForms(b: boolean): void
        setMatchWholeWord(b: boolean): void
        isAllLetterOrDigit(str: string): boolean
        clearAllFuzzyOptions(): void
        setCorrectHangulEndings(b: boolean): void
        isCorrectHangulEndings(): boolean
        getLanguageIDFarEast(): int
        setLanguageIDFarEast(id: int): void
        getLanguageIDOther(): int
        setLanguageIDOther(id: int): void
        setMatchAlefHamza(b: boolean): void
        isMatchDiacritics(): boolean
        setMatchDiacritics(b: boolean): void
        isMatchSoundsLike(): boolean
        setMatchSoundsLike(b: boolean): void
        setMatchWildcards(b: boolean): void
        setAllFuzzyOptions(): void
        getStyle(): any
        setStyle(style: any): void
        getLanguageID(): int
        setLanguageID(id: int): void
    }
    export interface FirstLetterException {
        getName(): string
        delete(): void
        getIndex(): int
    }
    export interface FirstLetterException {
        getName(): string
        delete(): void
        getIndex(): int
    }
    export interface FirstLetterExceptions {
        add(name: string): word.FirstLetterException
        item(index: any): word.FirstLetterException
        getCount(): int
        setFirstLetterExceptionName(firstLetterExceptionName: string): void
        getAllFirstLetterException(): any
        getFirstLetterExceptionName(): string
    }
    export interface FirstLetterExceptions {
        add(name: string): word.FirstLetterException
        item(index: any): word.FirstLetterException
        getCount(): int
        setFirstLetterExceptionName(firstLetterExceptionName: string): void
        getAllFirstLetterException(): any
        getFirstLetterExceptionName(): string
    }
    export interface Font {
        getName(): string
        apply(): void
        setName(name: string): void
        getSize(): double
        reset(): void
        grow(): void
        setSize(f: double): void
        setColor(l: int): void
        duplicate(): word.Font
        checkRange(): void
        getColor(): int
        setPosition(position: int): void
        getPosition(): int
        setColorIndex(l: int): void
        getColorIndex(): int
        setBorders(borders: word.Borders): void
        setShadow(st: int): void
        getBorders(): word.Borders
        getShading(): word.Shading
        initFont(): void
        setBold(l: int): void
        getBold(): int
        shrink(): void
        getFill(): word.FillFormat
        getMFontAttribute(): any
        setDoubleStrikeThrough(st: int): void
        setUnderlineColor(l: int): void
        getUnderlineColor(): int
        setDiacriticColor(l: int): void
        getDiacriticColor(): int
        getDoubleStrikeThrough(): int
        setSmallCaps(st: int): void
        setAllCaps(st: int): void
        setEmboss(st: int): void
        setEngrave(st: int): void
        setStrikeThrough(st: int): void
        setSubscript(st: int): void
        setSuperscript(st: int): void
        getKerning(): double
        setKerning(f: double): void
        getAllCaps(): int
        setAnimation(l: int): void
        getAnimation(): int
        setBoldBi(l: int): void
        getBoldBi(): int
        setColorIndexBi(l: int): void
        getColorIndexBi(): int
        getEmboss(): int
        setEmphasisMark(l: int): void
        getEmphasisMark(): int
        getEngrave(): int
        changeFontSize(increase: boolean): void
        setItalic(l: int): void
        getItalic(): int
        setItalicBi(l: int): void
        getItalicBi(): int
        setNameAscii(name: string): void
        getNameAscii(): string
        setNameBi(name: string): void
        getNameBi(): string
        setNameFarEast(name: string): void
        getNameFarEast(): string
        setNameOther(name: string): void
        getNameOther(): string
        setOutline(st: int): void
        getOutline(): int
        setScaling(l: int): void
        getScaling(): int
        getShadow(): int
        setSizeBi(f: double): void
        getSizeBi(): double
        getSmallCaps(): int
        setSpacing(f: double): void
        getSpacing(): double
        getStrikeThrough(): int
        getSubscript(): int
        getSuperscript(): int
        getTextColor(): word.ColorFormat
        setUnderline(l: int): void
        getUnderline(): int
        isDisableCharacterSpaceGrid(): boolean
        setDisableCharacterSpaceGrid(b: boolean): void
        setAsTemplateDefault(): void
        getHidden(): int
        setHidden(st: int): void
    }
    export interface Font {
        getName(): string
        apply(): void
        setName(name: string): void
        getSize(): double
        reset(): void
        grow(): void
        setSize(f: double): void
        setColor(l: int): void
        duplicate(): word.Font
        checkRange(): void
        getColor(): int
        setPosition(position: int): void
        getPosition(): int
        setColorIndex(l: int): void
        getColorIndex(): int
        setBorders(borders: word.Borders): void
        setShadow(st: int): void
        getBorders(): word.Borders
        getShading(): word.Shading
        initFont(): void
        setBold(l: int): void
        getBold(): int
        shrink(): void
        getFill(): word.FillFormat
        getMFontAttribute(): any
        setDoubleStrikeThrough(st: int): void
        setUnderlineColor(l: int): void
        getUnderlineColor(): int
        setDiacriticColor(l: int): void
        getDiacriticColor(): int
        getDoubleStrikeThrough(): int
        setSmallCaps(st: int): void
        setAllCaps(st: int): void
        setEmboss(st: int): void
        setEngrave(st: int): void
        setStrikeThrough(st: int): void
        setSubscript(st: int): void
        setSuperscript(st: int): void
        getKerning(): double
        setKerning(f: double): void
        getAllCaps(): int
        setAnimation(l: int): void
        getAnimation(): int
        setBoldBi(l: int): void
        getBoldBi(): int
        setColorIndexBi(l: int): void
        getColorIndexBi(): int
        getEmboss(): int
        setEmphasisMark(l: int): void
        getEmphasisMark(): int
        getEngrave(): int
        changeFontSize(increase: boolean): void
        setItalic(l: int): void
        getItalic(): int
        setItalicBi(l: int): void
        getItalicBi(): int
        setNameAscii(name: string): void
        getNameAscii(): string
        setNameBi(name: string): void
        getNameBi(): string
        setNameFarEast(name: string): void
        getNameFarEast(): string
        setNameOther(name: string): void
        getNameOther(): string
        setOutline(st: int): void
        getOutline(): int
        setScaling(l: int): void
        getScaling(): int
        getShadow(): int
        setSizeBi(f: double): void
        getSizeBi(): double
        getSmallCaps(): int
        setSpacing(f: double): void
        getSpacing(): double
        getStrikeThrough(): int
        getSubscript(): int
        getSuperscript(): int
        getTextColor(): word.ColorFormat
        setUnderline(l: int): void
        getUnderline(): int
        isDisableCharacterSpaceGrid(): boolean
        setDisableCharacterSpaceGrid(b: boolean): void
        setAsTemplateDefault(): void
        getHidden(): int
        setHidden(st: int): void
    }
    export interface FontNames {
        item(index: int): string
        getCount(): int
    }
    export interface FontNames {
        item(index: int): string
        getCount(): int
    }
    export interface Footnote {
        delete(): void
    }
    export interface Footnote {
        delete(): void
    }
    export interface FootnoteOptions {
        getLocation(): int
        setLocation(l: int): void
        setNumberStyle(l: int): void
        getNumberStyle(): int
        setStartingNumber(l: int): void
        getStartingNumber(): int
        setNumberingRule(l: int): void
        getNumberingRule(): int
    }
    export interface FootnoteOptions {
        getLocation(): int
        setLocation(l: int): void
        setNumberStyle(l: int): void
        getNumberStyle(): int
        setStartingNumber(l: int): void
        getStartingNumber(): int
        setNumberingRule(l: int): void
        getNumberingRule(): int
    }
    export interface Footnotes {
        add(range: word.Range, reference: any, text: any): word.Footnote
        getLocation(): int
        getSeparator(): word.Range
        convert(): void
        item(index: int): word.Footnote
        item(arg0: int): word.Note
        getCount(): int
        setLocation(location: int): void
        setNumberStyle(style: int): void
        getNumberStyle(): int
        resetContinuationSeparator(): void
        swapWithEndnotes(): void
        getContinuationNotice(): word.Range
        setStartingNumber(number: int): void
        getStartingNumber(): int
        getContinuationSeparator(): word.Range
        resetContinuationNotice(): void
        setNumberingRule(rule: int): void
        getNumberingRule(): int
        resetSeparator(): void
    }
    export interface Footnotes {
        add(range: word.Range, reference: any, text: any): word.Footnote
        getLocation(): int
        getSeparator(): word.Range
        convert(): void
        item(index: int): word.Footnote
        item(arg0: int): word.Note
        getCount(): int
        setLocation(location: int): void
        setNumberStyle(style: int): void
        getNumberStyle(): int
        resetContinuationSeparator(): void
        swapWithEndnotes(): void
        getContinuationNotice(): word.Range
        setStartingNumber(number: int): void
        getStartingNumber(): int
        getContinuationSeparator(): word.Range
        resetContinuationNotice(): void
        setNumberingRule(rule: int): void
        getNumberingRule(): int
        resetSeparator(): void
    }
    export interface FormField {
        getName(): string
        next(): word.FormField
        delete(): void
        setName(s: string): void
        copy(): void
        getType(): int
        previous(): word.FormField
        getResult(): string
        cut(): void
        setEnabled(b: boolean): void
        isEnabled(): boolean
        getRange(): word.Range
        select(): void
        getCheckBox(): word.CheckBox
        getDropDown(): word.DropDown
        getEntryMacro(): string
        setEntryMacro(s: string): void
        getExitMacro(): string
        setExitMacro(s: string): void
        getHelpText(): string
        setHelpText(s: string): void
        setOwnHelp(b: boolean): void
        isOwnHelp(): boolean
        setOwnStatus(b: boolean): void
        isOwnStatus(): boolean
        setResult(text: string): void
        getStatusText(): string
        setStatusText(s: string): void
        getTextInput(): word.TextInput
        setCalculateOnExit(b: boolean): void
        isCalculateOnExit(): boolean
    }
    export interface FormField {
        getName(): string
        next(): word.FormField
        delete(): void
        setName(s: string): void
        copy(): void
        getType(): int
        previous(): word.FormField
        getResult(): string
        cut(): void
        setEnabled(b: boolean): void
        isEnabled(): boolean
        getRange(): word.Range
        select(): void
        getCheckBox(): word.CheckBox
        getDropDown(): word.DropDown
        getEntryMacro(): string
        setEntryMacro(s: string): void
        getExitMacro(): string
        setExitMacro(s: string): void
        getHelpText(): string
        setHelpText(s: string): void
        setOwnHelp(b: boolean): void
        isOwnHelp(): boolean
        setOwnStatus(b: boolean): void
        isOwnStatus(): boolean
        setResult(text: string): void
        getStatusText(): string
        setStatusText(s: string): void
        getTextInput(): word.TextInput
        setCalculateOnExit(b: boolean): void
        isCalculateOnExit(): boolean
    }
    export interface FormFields {
        add(range: word.Range, type: int): word.FormField
        item(obj: any): word.FormField
        getCount(): int
        isShaded(): boolean
        setShaded(b: boolean): void
    }
    export interface FormFields {
        add(range: word.Range, type: int): word.FormField
        item(obj: any): word.FormField
        getCount(): int
        isShaded(): boolean
        setShaded(b: boolean): void
    }
    export interface FoundFiles {
        item(index: int): string
        getCount(): int
    }
    export interface FoundFiles {
        item(index: int): string
        getCount(): int
    }
    export interface Frame {
        apply(): void
        delete(): void
        copy(): void
        getWidth(): double
        setWidth(v: double): void
        getHeight(): double
        setHeight(v: double): void
        cut(): void
        setHeightRule(l: int): void
        getHeightRule(): int
        getRange(): word.Range
        select(): void
        getBorders(): word.Borders
        getShading(): word.Shading
        setLockAnchor(b: boolean): void
        isLockAnchor(): boolean
        setTextWrap(b: boolean): void
        isTextWrap(): boolean
        getWidthRule(): int
        setWidthRule(l: int): void
        setHorizontalDistanceFromText(f: double): void
        getHorizontalDistanceFromText(): double
        setVerticalDistanceFromText(f: double): void
        getVerticalDistanceFromText(): double
        getRelativeHorizontalPosition(): int
        setRelativeHorizontalPosition(l: int): void
        getRelativeVerticalPosition(): int
        setRelativeVerticalPosition(l: int): void
        setHorizontalPosition(f: double): void
        getHorizontalPosition(): double
        setVerticalPosition(f: double): void
        getVerticalPosition(): double
    }
    export interface Frame {
        apply(): void
        delete(): void
        copy(): void
        getWidth(): double
        setWidth(v: double): void
        getHeight(): double
        setHeight(v: double): void
        cut(): void
        setHeightRule(l: int): void
        getHeightRule(): int
        getRange(): word.Range
        select(): void
        getBorders(): word.Borders
        getShading(): word.Shading
        setLockAnchor(b: boolean): void
        isLockAnchor(): boolean
        setTextWrap(b: boolean): void
        isTextWrap(): boolean
        getWidthRule(): int
        setWidthRule(l: int): void
        setHorizontalDistanceFromText(f: double): void
        getHorizontalDistanceFromText(): double
        setVerticalDistanceFromText(f: double): void
        getVerticalDistanceFromText(): double
        getRelativeHorizontalPosition(): int
        setRelativeHorizontalPosition(l: int): void
        getRelativeVerticalPosition(): int
        setRelativeVerticalPosition(l: int): void
        setHorizontalPosition(f: double): void
        getHorizontalPosition(): double
        setVerticalPosition(f: double): void
        getVerticalPosition(): double
    }
    export interface Frames {
        add(range: word.Range): word.Frame
        delete(): void
        item(index: int): word.Frame
        getCount(): int
    }
    export interface Frames {
        add(range: word.Range): word.Frame
        delete(): void
        item(index: int): word.Frame
        getCount(): int
    }
    export interface Frameset {
        delete(): void
        getType(): int
        getWidth(): int
        setWidth(l: int): void
        getHeight(): int
        setHeight(l: int): void
        addNewFrame(where: int): word.Frameset
        setFrameName(s: string): void
        getFrameName(): string
        isFrameResizable(): boolean
        getHeightType(): int
        setHeightType(l: int): void
        getWidthType(): int
        setWidthType(l: int): void
        getChildFramesetCount(): int
        getChildFramesetItem(index: int): word.Frameset
        setFrameDefaultURL(url: string): void
        getFrameDefaultURL(): string
        setFrameDisplayBorders(b: boolean): void
        isFrameDisplayBorders(): boolean
        setFrameLinkToFile(b: boolean): void
        isFrameLinkToFile(): boolean
        setFrameResizable(b: boolean): void
        getFrameScrollbarType(): int
        setFrameScrollbarType(type: int): void
        getFramesetBorderColor(): int
        setFramesetBorderColor(type: int): void
        getFramesetBorderWidth(): double
        setFramesetBorderWidth(f: double): void
        getParentFrameset(): word.Frameset
    }
    export interface Frameset {
        delete(): void
        getType(): int
        getWidth(): int
        setWidth(l: int): void
        getHeight(): int
        setHeight(l: int): void
        addNewFrame(where: int): word.Frameset
        setFrameName(s: string): void
        getFrameName(): string
        isFrameResizable(): boolean
        getHeightType(): int
        setHeightType(l: int): void
        getWidthType(): int
        setWidthType(l: int): void
        getChildFramesetCount(): int
        getChildFramesetItem(index: int): word.Frameset
        setFrameDefaultURL(url: string): void
        getFrameDefaultURL(): string
        setFrameDisplayBorders(b: boolean): void
        isFrameDisplayBorders(): boolean
        setFrameLinkToFile(b: boolean): void
        isFrameLinkToFile(): boolean
        setFrameResizable(b: boolean): void
        getFrameScrollbarType(): int
        setFrameScrollbarType(type: int): void
        getFramesetBorderColor(): int
        setFramesetBorderColor(type: int): void
        getFramesetBorderWidth(): double
        setFramesetBorderWidth(f: double): void
        getParentFrameset(): word.Frameset
    }
    export interface FreeformBuilder {
        addNodes(segmentType: int, editingType: int, X1: double, Y1: double, X2: double, Y2: double, X3: double, Y3: double): void
        convertToShape(anchor: any): word.Shape
        addNode(x: double, y: double): void
    }
    export interface FreeformBuilder {
        addNodes(segmentType: int, editingType: int, X1: double, Y1: double, X2: double, Y2: double, X3: double, Y3: double): void
        convertToShape(anchor: any): word.Shape
        addNode(x: double, y: double): void
    }
    export interface Global {
        createDocument(): word.Document
        loadDocument(pathStr: string, readOnly: boolean): word.Document
    }
    export interface Global {
        createDocument(): word.Document
        loadDocument(pathStr: string, readOnly: boolean): word.Document
    }
    export interface GroupShapes {
        range(index: any): word.ShapeRange
        range(index: string): word.ShapeRange
        item(index: any): word.Shape
        getCount(): int
    }
    export interface GroupShapes {
        range(index: any): word.ShapeRange
        range(index: string): word.ShapeRange
        item(index: any): word.Shape
        getCount(): int
    }
    export interface HangulAndAlphabetException {
        getName(): string
        delete(): void
        getIndex(): int
    }
    export interface HangulAndAlphabetException {
        getName(): string
        delete(): void
        getIndex(): int
    }
    export interface HangulAndAlphabetExceptions {
        add(name: string): word.HangulAndAlphabetException
        item(index: any): word.HangulAndAlphabetException
        getCount(): int
    }
    export interface HangulAndAlphabetExceptions {
        add(name: string): word.HangulAndAlphabetException
        item(index: any): word.HangulAndAlphabetException
        getCount(): int
    }
    export interface HangulHanjaConversionDictionaries {
        add(fileName: string): word.Dictionary
        item(index: int): word.Dictionary
        getCount(): int
        setActiveCustomDictionary(dic: word.Dictionary): void
        getActiveCustomDictionary(): word.Dictionary
        getBuiltinDictionary(): word.Dictionary
        clearAll(): void
        getMaximum(): int
    }
    export interface HangulHanjaConversionDictionaries {
        add(fileName: string): word.Dictionary
        item(index: int): word.Dictionary
        getCount(): int
        setActiveCustomDictionary(dic: word.Dictionary): void
        getActiveCustomDictionary(): word.Dictionary
        getBuiltinDictionary(): word.Dictionary
        clearAll(): void
        getMaximum(): int
    }
    export interface HeaderFooter {
        initIndex(doc: word.Document, section: word.Section): int
        getIndex(): int
        getRange(): word.Range
        isExists(): boolean
        isHeader(): boolean
        setExists(b: boolean): void
        isLinkToPrevious(): boolean
        getPageNumbers(): word.PageNumbers
        setLinkToPrevious(b: boolean): void
        getShapes(): word.Shapes
    }
    export interface HeaderFooter {
        initIndex(doc: word.Document, section: word.Section): int
        getIndex(): int
        getRange(): word.Range
        isExists(): boolean
        isHeader(): boolean
        setExists(b: boolean): void
        isLinkToPrevious(): boolean
        getPageNumbers(): word.PageNumbers
        setLinkToPrevious(b: boolean): void
        getShapes(): word.Shapes
    }
    export interface HeadersFooters {
        item(index: int): word.HeaderFooter
        getCount(): int
    }
    export interface HeadersFooters {
        item(index: int): word.HeaderFooter
        getCount(): int
    }
    export interface HeadingStyle {
        delete(): void
        setLevel(level: int): void
        getStyle(): any
        setStyle(style: any): void
        getLevel(): int
    }
    export interface HeadingStyle {
        delete(): void
        setLevel(level: int): void
        getStyle(): any
        setStyle(style: any): void
        getLevel(): int
    }
    export interface HeadingStyles {
        add(style: any, level: int): word.HeadingStyle
        item(index: int): word.HeadingStyle
        getCount(): int
    }
    export interface HeadingStyles {
        add(style: any, level: int): word.HeadingStyle
        item(index: int): word.HeadingStyle
        getCount(): int
    }
    export interface HorizontalLineFormat {
        getWidthType(): int
        setWidthType(l: int): void
        getAlignment(): int
        setAlignment(l: int): void
        isNoShade(): boolean
        setNoShade(b: boolean): void
        getPercentWidth(): double
        setPercentWidth(b: double): void
    }
    export interface HorizontalLineFormat {
        getWidthType(): int
        setWidthType(l: int): void
        getAlignment(): int
        setAlignment(l: int): void
        isNoShade(): boolean
        setNoShade(b: boolean): void
        getPercentWidth(): double
        setPercentWidth(b: double): void
    }
    export interface HTMLDivision {
        delete(): void
        getRange(): word.Range
        getBorders(): word.Borders
        getLeftIndent(): double
        setLeftIndent(f: double): void
        getRightIndent(): double
        setRightIndent(f: double): void
        getSpaceAfter(): double
        setSpaceAfter(f: double): void
        getSpaceBefore(): double
        setSpaceBefore(f: double): void
        HTMLDivisionParent(levelsUp: int): word.HTMLDivision
        getHTMLDivisions(): word.HTMLDivisions
    }
    export interface HTMLDivision {
        delete(): void
        getRange(): word.Range
        getBorders(): word.Borders
        getLeftIndent(): double
        setLeftIndent(f: double): void
        getRightIndent(): double
        setRightIndent(f: double): void
        getSpaceAfter(): double
        setSpaceAfter(f: double): void
        getSpaceBefore(): double
        setSpaceBefore(f: double): void
        HTMLDivisionParent(levelsUp: int): word.HTMLDivision
        getHTMLDivisions(): word.HTMLDivisions
    }
    export interface HTMLDivisions {
        add(range: word.Range): word.HTMLDivision
        item(index: int): word.HTMLDivision
        getCount(): int
        getNestingLevel(): int
    }
    export interface HTMLDivisions {
        add(range: word.Range): word.HTMLDivision
        item(index: int): word.HTMLDivision
        getCount(): int
        getNestingLevel(): int
    }
    export interface Hyperlink {
        getAddress(): string
        getName(): string
        apply(): void
        init(): void
        delete(): void
        getType(): int
        getTarget(): string
        setTarget(s: string): void
        getShape(): word.Shape
        getRange(): word.Range
        follow(NewWindow: any, AddHistory: any, ExtraInfo: any, Method: any, HeaderInfo: any): void
        setTextToDisplay(s: string): void
        setAddress(s: string): void
        setEmailSubject(s: string): void
        getEmailSubject(): string
        setScreenTip(s: string): void
        getScreenTip(): string
        setSubAddress(s: string): void
        getSubAddress(): string
        getTextToDisplay(): string
        createNewDocument(fileName: string, EditNow: boolean, Overwrite: boolean): void
        setExtraInfoRequired(b: boolean): void
        isExtraInfoRequired(): boolean
        addToFavorites(): void
    }
    export interface Hyperlink {
        getAddress(): string
        getName(): string
        apply(): void
        init(): void
        delete(): void
        getType(): int
        getTarget(): string
        setTarget(s: string): void
        getShape(): word.Shape
        getRange(): word.Range
        follow(NewWindow: any, AddHistory: any, ExtraInfo: any, Method: any, HeaderInfo: any): void
        setTextToDisplay(s: string): void
        setAddress(s: string): void
        setEmailSubject(s: string): void
        getEmailSubject(): string
        setScreenTip(s: string): void
        getScreenTip(): string
        setSubAddress(s: string): void
        getSubAddress(): string
        getTextToDisplay(): string
        createNewDocument(fileName: string, EditNow: boolean, Overwrite: boolean): void
        setExtraInfoRequired(b: boolean): void
        isExtraInfoRequired(): boolean
        addToFavorites(): void
    }
    export interface Hyperlinks {
        add(anchorObj: any, addressObj: any, subAddressObj: any, screenTipObj: any, textToDisplayObj: any, targetObj: any): word.Hyperlink
        item(index: any): word.Hyperlink
        getCount(): int
    }
    export interface Hyperlinks {
        add(anchorObj: any, addressObj: any, subAddressObj: any, screenTipObj: any, textToDisplayObj: any, targetObj: any): word.Hyperlink
        item(index: any): word.Hyperlink
        getCount(): int
    }
    export interface IMTable {
        getMRange(): any
        getMTable(): any
    }
    export interface IMTable {
        getMRange(): any
        getMTable(): any
    }
    export interface Index {
        update(): void
        delete(): void
        getType(): int
        setType(l: int): void
        getRange(): word.Range
        getFilter(): int
        setFilter(l: int): void
        getIndexLanguage(): int
        setIndexLanguage(l: int): void
        getSortBy(): int
        setSortBy(l: int): void
        getTabLeader(): int
        setTabLeader(l: int): void
        setAccentedLetters(b: boolean): void
        isAccentedLetters(): boolean
        getHeadingSeparator(): int
        setHeadingSeparator(l: int): void
        getNumberOfColumns(): int
        setNumberOfColumns(l: int): void
        setRightAlignPageNumbers(b: boolean): void
        isRightAlignPageNumbers(): boolean
    }
    export interface Index {
        update(): void
        delete(): void
        getType(): int
        setType(l: int): void
        getRange(): word.Range
        getFilter(): int
        setFilter(l: int): void
        getIndexLanguage(): int
        setIndexLanguage(l: int): void
        getSortBy(): int
        setSortBy(l: int): void
        getTabLeader(): int
        setTabLeader(l: int): void
        setAccentedLetters(b: boolean): void
        isAccentedLetters(): boolean
        getHeadingSeparator(): int
        setHeadingSeparator(l: int): void
        getNumberOfColumns(): int
        setNumberOfColumns(l: int): void
        setRightAlignPageNumbers(b: boolean): void
        isRightAlignPageNumbers(): boolean
    }
    export interface Indexes {
        getCount(): int
        Item(index: int): word.Index
        getFormat(): int
        Add(range: word.Range, headingSeparator: int, rightAlignPageNumbers: boolean, type: int, numberOfColumns: int, accentedLetters: int, sortBy: int, indexLanguage: int): word.Index
        setFormat(l: int): void
        autoMarkEntries(concordanceFileName: string): void
        MarkAllEntries(range: word.Range, entry: string, entryAutoText: string, crossReference: string, crossReferenceAutoText: string, bookmarkName: string, bold: boolean, italic: boolean): void
        MarkEntry(range: word.Range, entry: string, entryAutoText: string, crossReference: string, crossReferenceAutoText: string, bookmarkName: string, bold: boolean, italic: boolean, Reading: boolean): word.Field
    }
    export interface Indexes {
        getCount(): int
        Item(index: int): word.Index
        getFormat(): int
        Add(range: word.Range, headingSeparator: int, rightAlignPageNumbers: boolean, type: int, numberOfColumns: int, accentedLetters: int, sortBy: int, indexLanguage: int): word.Index
        setFormat(l: int): void
        autoMarkEntries(concordanceFileName: string): void
        MarkAllEntries(range: word.Range, entry: string, entryAutoText: string, crossReference: string, crossReferenceAutoText: string, bookmarkName: string, bold: boolean, italic: boolean): void
        MarkEntry(range: word.Range, entry: string, entryAutoText: string, crossReference: string, crossReferenceAutoText: string, bookmarkName: string, bold: boolean, italic: boolean, Reading: boolean): word.Field
    }
    export interface InlineShape {
        getName(): string
        getField(): word.Field
        delete(): void
        getType(): int
        reset(): void
        getScript(): office.Script
        getShape(): any
        setTitle(title: string): void
        getTitle(): string
        getWidth(): double
        setWidth(f: double): void
        activate(): void
        getHeight(): double
        setHeight(f: double): void
        saveAs(savePath: string): void
        getRange(): word.Range
        select(): void
        getBorders(): word.Borders
        getFill(): word.FillFormat
        getLine(): word.LineFormat
        getOLEFormat(): word.OLEFormat
        getHyperlink(): word.Hyperlink
        getShadow(): word.ShadowFormat
        convertToShape(): word.Shape
        isPictureBullet(): boolean
        getHasChart(): int
        getHasSmartArt(): int
        getPictureFormat(): word.PictureFormat
        setScaleHeight(f: double): void
        getScaleHeight(): double
        setScaleWidth(f: double): void
        getScaleWidth(): double
        getTextEffect(): word.TextEffectFormat
        getGroupItems(): word.GroupShapes
        setAlternativeText(s: string): void
        getAlternativeText(): string
        getHorizontalLineFormat(): word.HorizontalLineFormat
        setLockAspectRatio(l: int): void
        getLockAspectRatio(): int
        getLinkFormat(): word.LinkFormat
    }
    export interface InlineShape {
        getName(): string
        getField(): word.Field
        delete(): void
        getType(): int
        reset(): void
        getScript(): office.Script
        getShape(): any
        setTitle(title: string): void
        getTitle(): string
        getWidth(): double
        setWidth(f: double): void
        activate(): void
        getHeight(): double
        setHeight(f: double): void
        saveAs(savePath: string): void
        getRange(): word.Range
        select(): void
        getBorders(): word.Borders
        getFill(): word.FillFormat
        getLine(): word.LineFormat
        getOLEFormat(): word.OLEFormat
        getHyperlink(): word.Hyperlink
        getShadow(): word.ShadowFormat
        convertToShape(): word.Shape
        isPictureBullet(): boolean
        getHasChart(): int
        getHasSmartArt(): int
        getPictureFormat(): word.PictureFormat
        setScaleHeight(f: double): void
        getScaleHeight(): double
        setScaleWidth(f: double): void
        getScaleWidth(): double
        getTextEffect(): word.TextEffectFormat
        getGroupItems(): word.GroupShapes
        setAlternativeText(s: string): void
        getAlternativeText(): string
        getHorizontalLineFormat(): word.HorizontalLineFormat
        setLockAspectRatio(l: int): void
        getLockAspectRatio(): int
        getLinkFormat(): word.LinkFormat
    }
    export interface InlineShapes {
        checkRange(): void
        item(index: int): word.InlineShape
        getCount(): int
        addPicture(fileName: string, linkToFileObj: any, saveWithDocumentObj: any, rangeObj: any): word.InlineShape
        addHorizontalLineStandard(range: any): word.InlineShape
        addOLEControl(classType: any, range: any): word.InlineShape
        addOLEObject(classType: any, fileName: any, linkToFile: any, displayAsIcon: any, iconFileName: any, iconIndex: any, iconLabel: any, range: any): word.InlineShape
        addPictureBullet(fileName: string, range: any): word.InlineShape
        newLineShape(range: word.Range): word.InlineShape
        addHorizontalLine(fileName: string, range: any): word.InlineShape
    }
    export interface InlineShapes {
        checkRange(): void
        item(index: int): word.InlineShape
        getCount(): int
        addPicture(fileName: string, linkToFileObj: any, saveWithDocumentObj: any, rangeObj: any): word.InlineShape
        addHorizontalLineStandard(range: any): word.InlineShape
        addOLEControl(classType: any, range: any): word.InlineShape
        addOLEObject(classType: any, fileName: any, linkToFile: any, displayAsIcon: any, iconFileName: any, iconIndex: any, iconLabel: any, range: any): word.InlineShape
        addPictureBullet(fileName: string, range: any): word.InlineShape
        newLineShape(range: word.Range): word.InlineShape
        addHorizontalLine(fileName: string, range: any): word.InlineShape
    }
    export interface KeyBinding {
        clear(): void
        execute(): void
        getContext(): any
        rebind(keyCategory: int, command: string, commandParameter: string): void
        disable(): void
        getKeyCode(): int
        getCommand(): string
        getKeyCategory(): int
        getKeyCode2(): int
        getKeyString(): string
        getProtected(): boolean
        getCommandParameter(): string
    }
    export interface KeyBinding {
        clear(): void
        execute(): void
        getContext(): any
        rebind(keyCategory: int, command: string, commandParameter: string): void
        disable(): void
        getKeyCode(): int
        getCommand(): string
        getKeyCategory(): int
        getKeyCode2(): int
        getKeyString(): string
        getProtected(): boolean
        getCommandParameter(): string
    }
    export interface KeyBindings {
        add(keyCategory: int, command: string, keyCode: int, keyCode2: any, commandParameter: any): word.KeyBinding
        getContext(): any
        Key(keyCode: int, keyCode2: int): word.KeyBinding
        item(index: int): word.KeyBinding
        getCount(): int
        clearAll(): void
    }
    export interface KeyBindings {
        add(keyCategory: int, command: string, keyCode: int, keyCode2: any, commandParameter: any): word.KeyBinding
        getContext(): any
        Key(keyCode: int, keyCode2: int): word.KeyBinding
        item(index: int): word.KeyBinding
        getCount(): int
        clearAll(): void
    }
    export interface KeysBoundTo {
        getContext(): any
        Key(keyCode: int, KeyCode2: int): word.KeyBinding
        item(index: int): word.KeyBinding
        getCount(): int
        getCommand(): string
        getKeyCategory(): int
        getCommandParameter(): any
    }
    export interface KeysBoundTo {
        getContext(): any
        Key(keyCode: int, KeyCode2: int): word.KeyBinding
        item(index: int): word.KeyBinding
        getCount(): int
        getCommand(): string
        getKeyCategory(): int
        getCommandParameter(): any
    }
    export interface Language {
        getName(): string
        getID(): int
        getNameLocal(): string
        getActiveGrammarDictionary(): word.Dictionary
        getActiveHyphenationDictionary(): word.Dictionary
        getActiveSpellingDictionary(): word.Dictionary
        getActiveThesaurusDictionary(): word.Dictionary
        getSpellingDictionaryType(): int
        setSpellingDictionaryType(spellingDictionaryType: int): void
        getDefaultWritingStyle(): string
        setDefaultWritingStyle(writingStyle: string): void
        getWritingStyleList(): string[]
    }
    export interface Language {
        getName(): string
        getID(): int
        getNameLocal(): string
        getActiveGrammarDictionary(): word.Dictionary
        getActiveHyphenationDictionary(): word.Dictionary
        getActiveSpellingDictionary(): word.Dictionary
        getActiveThesaurusDictionary(): word.Dictionary
        getSpellingDictionaryType(): int
        setSpellingDictionaryType(spellingDictionaryType: int): void
        getDefaultWritingStyle(): string
        setDefaultWritingStyle(writingStyle: string): void
        getWritingStyleList(): string[]
    }
    export interface Languages {
        item(index: any): word.Language
        getCount(): int
    }
    export interface Languages {
        item(index: any): word.Language
        getCount(): int
    }
    export interface LetterContent {
        getAttentionLine(): string
        setAttentionLine(attentionLine: string): void
        getCCList(): string
        setCCList(cclist: string): void
        getClosing(): string
        setClosing(closing: string): void
        getDateFormat(): string
        setDateFormat(dateFormat: string): void
        getDuplicate(): word.LetterContent
        isInfoBlock(): boolean
        setInfoBlock(infoBlock: boolean): void
        isLetterhead(): boolean
        setLetterhead(letterhead: boolean): void
        getLetterStyle(): int
        setLetterStyle(letterStyle: int): void
        getEnclosureNumber(): int
        setEnclosureNumber(enclosureNumber: int): void
        isIncludeHeaderFooter(): boolean
        setIncludeHeaderFooter(includeHeaderFooter: boolean): void
        getLetterheadLocation(): int
        setLetterheadLocation(letterheadLocation: int): void
        getLetterheadSize(): double
        setLetterheadSize(letterheadSize: double): void
    }
    export interface LetterContent {
        getAttentionLine(): string
        setAttentionLine(attentionLine: string): void
        getCCList(): string
        setCCList(cclist: string): void
        getClosing(): string
        setClosing(closing: string): void
        getDateFormat(): string
        setDateFormat(dateFormat: string): void
        getDuplicate(): word.LetterContent
        isInfoBlock(): boolean
        setInfoBlock(infoBlock: boolean): void
        isLetterhead(): boolean
        setLetterhead(letterhead: boolean): void
        getLetterStyle(): int
        setLetterStyle(letterStyle: int): void
        getEnclosureNumber(): int
        setEnclosureNumber(enclosureNumber: int): void
        isIncludeHeaderFooter(): boolean
        setIncludeHeaderFooter(includeHeaderFooter: boolean): void
        getLetterheadLocation(): int
        setLetterheadLocation(letterheadLocation: int): void
        getLetterheadSize(): double
        setLetterheadSize(letterheadSize: double): void
    }
    export interface Line {
        getWidth(): int
        getLeft(): int
        getTop(): int
        setWidth(value: int): void
        getHeight(): int
        setHeight(value: int): void
        getRang(): word.Range
        getLineType(): int
        getRectangles(): word.Rectangles
    }
    export interface Line {
        getWidth(): int
        getLeft(): int
        getTop(): int
        setWidth(value: int): void
        getHeight(): int
        setHeight(value: int): void
        getRang(): word.Range
        getLineType(): int
        getRectangles(): word.Rectangles
    }
    export interface LineFormat {
        setVisible(visible: int): void
        getBackColor(): word.ColorFormat
        getForeColor(): word.ColorFormat
        getPattern(): int
        getTransparency(): double
        setTransparency(pTransparency: double): void
        getVisible(): int
        setForeColor(foreColor: word.ColorFormat): void
        setBackColor(backColor: word.ColorFormat): void
        getDashStyle(): int
        setDashStyle(dashStyle: int): void
        getInsetPen(): int
        setInsetPen(insetPen: int): void
        setPattern(pattern: int): void
        getWeight(): double
        setWeight(weight: double): void
        getStyle(): int
        setStyle(style: int): void
        getBeginArrowheadLength(): int
        setBeginArrowheadLength(length: int): void
        getBeginArrowheadStyle(): int
        transArrowStyleToMso(type: int): int
        setBeginArrowheadStyle(type: int): void
        transArrowStyleToYzo(type: int): int
        getBeginArrowheadWidth(): int
        setBeginArrowheadWidth(width: int): void
        getEndArrowheadLength(): int
        setEndArrowheadLength(length: int): void
        getEndArrowheadStyle(): int
        setEndArrowheadStyle(type: int): void
        getEndArrowheadWidth(): int
        setEndArrowheadWidth(width: int): void
    }
    export interface LineFormat {
        setVisible(visible: int): void
        getBackColor(): word.ColorFormat
        getForeColor(): word.ColorFormat
        getPattern(): int
        getTransparency(): double
        setTransparency(pTransparency: double): void
        getVisible(): int
        setForeColor(foreColor: word.ColorFormat): void
        setBackColor(backColor: word.ColorFormat): void
        getDashStyle(): int
        setDashStyle(dashStyle: int): void
        getInsetPen(): int
        setInsetPen(insetPen: int): void
        setPattern(pattern: int): void
        getWeight(): double
        setWeight(weight: double): void
        getStyle(): int
        setStyle(style: int): void
        getBeginArrowheadLength(): int
        setBeginArrowheadLength(length: int): void
        getBeginArrowheadStyle(): int
        transArrowStyleToMso(type: int): int
        setBeginArrowheadStyle(type: int): void
        transArrowStyleToYzo(type: int): int
        getBeginArrowheadWidth(): int
        setBeginArrowheadWidth(width: int): void
        getEndArrowheadLength(): int
        setEndArrowheadLength(length: int): void
        getEndArrowheadStyle(): int
        setEndArrowheadStyle(type: int): void
        getEndArrowheadWidth(): int
        setEndArrowheadWidth(width: int): void
    }
    export interface LineNumbering {
        getActive(): int
        setActive(active: int): void
        getCountBy(): int
        setCountBy(countBy: int): void
        getRestartMode(): int
        setRestartMode(restartMode: int): void
        setDistanceFromText(distanceFromText: double): void
        getDistanceFromText(): double
        setStartingNumber(startingNumber: int): void
        getStartingNumber(): int
    }
    export interface LineNumbering {
        getActive(): int
        setActive(active: int): void
        getCountBy(): int
        setCountBy(countBy: int): void
        getRestartMode(): int
        setRestartMode(restartMode: int): void
        setDistanceFromText(distanceFromText: double): void
        getDistanceFromText(): double
        setStartingNumber(startingNumber: int): void
        getStartingNumber(): int
    }
    export interface Lines {
        item(index: int): word.Line
        getCount(): int
    }
    export interface Lines {
        item(index: int): word.Line
        getCount(): int
    }
    export interface LinkFormat {
        update(): void
        getType(): int
        isLocked(): boolean
        isAutoUpdate(): boolean
        setAutoUpdate(autoUpdate: boolean): void
        getSourceName(): string
        getSourcePath(): string
        breakLink(): void
        isSavePictureWithDocument(): boolean
        setSavePictureWithDocument(savePictureWithDocument: boolean): void
        setLocked(locked: boolean): void
        getSourceFullName(): string
        setSourceFullName(sourceFullName: string): void
    }
    export interface LinkFormat {
        update(): void
        getType(): int
        isLocked(): boolean
        isAutoUpdate(): boolean
        setAutoUpdate(autoUpdate: boolean): void
        getSourceName(): string
        getSourcePath(): string
        breakLink(): void
        isSavePictureWithDocument(): boolean
        setSavePictureWithDocument(savePictureWithDocument: boolean): void
        setLocked(locked: boolean): void
        getSourceFullName(): string
        setSourceFullName(sourceFullName: string): void
    }
    export interface List {
        getRange(): word.Range
        getStyleName(): string
        convertNumbersToText(numberType: any): void
        countNumberedItems(numberType: any, level: any): int
        removeNumbers(numberType: any): void
        getListParagraphs(): word.ListParagraphs
        isSingleListTemplate(): boolean
        applyListTemplate(listTemplate: word.ListTemplate, continuePreviousList: any, defaultListBehavior: any): void
        canContinuePreviousList(listTemplate: word.ListTemplate): int
    }
    export interface List {
        getRange(): word.Range
        getStyleName(): string
        convertNumbersToText(numberType: any): void
        countNumberedItems(numberType: any, level: any): int
        removeNumbers(numberType: any): void
        getListParagraphs(): word.ListParagraphs
        isSingleListTemplate(): boolean
        applyListTemplate(listTemplate: word.ListTemplate, continuePreviousList: any, defaultListBehavior: any): void
        canContinuePreviousList(listTemplate: word.ListTemplate): int
    }
    export interface ListEntries {
        add(name: string, index: any): word.ListEntry
        clear(): void
        delete(entry: word.ListEntry): void
        item(index: any): word.ListEntry
        getIndex(entry: word.ListEntry): int
        getCount(): int
    }
    export interface ListEntries {
        add(name: string, index: any): word.ListEntry
        clear(): void
        delete(entry: word.ListEntry): void
        item(index: any): word.ListEntry
        getIndex(entry: word.ListEntry): int
        getCount(): int
    }
    export interface ListEntry {
        getName(): string
        delete(): void
        setName(name: string): void
        getIndex(): int
    }
    export interface ListEntry {
        getName(): string
        delete(): void
        setName(name: string): void
        getIndex(): int
    }
    export interface ListFormat {
        convertNumbersToText(numberType: any): void
        countNumberedItems(numberType: any, level: any): int
        getListString(): string
        getListTemplate(): word.ListTemplate
        getListValue(): int
        getListType(): int
        isSingleList(): boolean
        listIndent(): void
        listOutdent(): void
        applyOutlineNumberDefault(defaultListBehavior: any): void
        getList(): word.List
        removeNumbers(numberType: any): void
        isSingleListTemplate(): boolean
        applyListTemplate(listTemplate: word.ListTemplate, continuePreviousList: any, applyTo: any, defaultListBehavior: any): void
        canContinuePreviousList(listTemplate: word.ListTemplate): int
        getListLevelNumber(): int
        setListLevelNumber(listLevelNumber: int): void
        getListPictureBullet(): word.InlineShape
        applyBulletDefault(defaultListBehavior: any): void
        applyNumberDefault(defaultListBehavior: any): void
    }
    export interface ListFormat {
        convertNumbersToText(numberType: any): void
        countNumberedItems(numberType: any, level: any): int
        getListString(): string
        getListTemplate(): word.ListTemplate
        getListValue(): int
        getListType(): int
        isSingleList(): boolean
        listIndent(): void
        listOutdent(): void
        applyOutlineNumberDefault(defaultListBehavior: any): void
        getList(): word.List
        removeNumbers(numberType: any): void
        isSingleListTemplate(): boolean
        applyListTemplate(listTemplate: word.ListTemplate, continuePreviousList: any, applyTo: any, defaultListBehavior: any): void
        canContinuePreviousList(listTemplate: word.ListTemplate): int
        getListLevelNumber(): int
        setListLevelNumber(listLevelNumber: int): void
        getListPictureBullet(): word.InlineShape
        applyBulletDefault(defaultListBehavior: any): void
        applyNumberDefault(defaultListBehavior: any): void
    }
    export interface ListGalleries {
        item(index: int): word.ListGallery
        getCount(): int
    }
    export interface ListGalleries {
        item(index: int): word.ListGallery
        getCount(): int
    }
    export interface ListGallery {
        reset(index: int): void
        isModified(index: int): boolean
        getListTemplates(): word.ListTemplates
    }
    export interface ListGallery {
        reset(index: int): void
        isModified(index: int): boolean
        getListTemplates(): word.ListTemplates
    }
    export interface ListLevel {
        getFont(): word.Font
        getIndex(): int
        setFont(font: word.Font): void
        setNumberStyle(numberStyle: int): void
        getNumberStyle(): int
        getAlignment(): int
        setAlignment(alignment: int): void
        getLinkedStyle(): string
        setLinkedStyle(linkedStyle: string): void
        getNumberFormat(): string
        setNumberFormat(numberFormat: string): void
        getPictureBullet(): word.InlineShape
        isResetOnHigher(): boolean
        setResetOnHigher(resetOnHigher: int): void
        getStartAt(): int
        setStartAt(startAt: int): void
        getTabPosition(): double
        setTabPosition(tabPosition: double): void
        getNumberPosition(): double
        setNumberPosition(numberPosition: double): void
        getTrailingCharacter(): int
        setTrailingCharacter(trailingCharacter: int): void
        applyPictureBullet(fileName: string): word.InlineShape
        getTextPosition(): double
        setTextPosition(textPosition: double): void
    }
    export interface ListLevel {
        getFont(): word.Font
        getIndex(): int
        setFont(font: word.Font): void
        setNumberStyle(numberStyle: int): void
        getNumberStyle(): int
        getAlignment(): int
        setAlignment(alignment: int): void
        getLinkedStyle(): string
        setLinkedStyle(linkedStyle: string): void
        getNumberFormat(): string
        setNumberFormat(numberFormat: string): void
        getPictureBullet(): word.InlineShape
        isResetOnHigher(): boolean
        setResetOnHigher(resetOnHigher: int): void
        getStartAt(): int
        setStartAt(startAt: int): void
        getTabPosition(): double
        setTabPosition(tabPosition: double): void
        getNumberPosition(): double
        setNumberPosition(numberPosition: double): void
        getTrailingCharacter(): int
        setTrailingCharacter(trailingCharacter: int): void
        applyPictureBullet(fileName: string): word.InlineShape
        getTextPosition(): double
        setTextPosition(textPosition: double): void
    }
    export interface ListLevels {
        item(index: int): word.ListLevel
        getCount(): int
    }
    export interface ListLevels {
        item(index: int): word.ListLevel
        getCount(): int
    }
    export interface ListParagraphs {
        item(index: int): word.Paragraph
        getCount(): int
    }
    export interface ListParagraphs {
        item(index: int): word.Paragraph
        getCount(): int
    }
    export interface Lists {
        item(index: int): word.List
        getCount(): int
    }
    export interface Lists {
        item(index: int): word.List
        getCount(): int
    }
    export interface ListTemplate {
        getName(): string
        setName(name: string): void
        convert(level: int): word.ListTemplate
        isOutlineNumbered(): boolean
        setOutlineNumbered(outlineNumbered: boolean): void
        getListLevels(): word.ListLevels
    }
    export interface ListTemplate {
        getName(): string
        setName(name: string): void
        convert(level: int): word.ListTemplate
        isOutlineNumbered(): boolean
        setOutlineNumbered(outlineNumbered: boolean): void
        getListLevels(): word.ListLevels
    }
    export interface ListTemplates {
        add(outlineNumbered: boolean, name: string): word.ListTemplate
        item(index: any): word.ListTemplate
        getCount(): int
    }
    export interface ListTemplates {
        add(outlineNumbered: boolean, name: string): word.ListTemplate
        item(index: any): word.ListTemplate
        getCount(): int
    }
    export interface Mailer {
        getBCCRecipients(): any
        setBCCRecipients(bCCRecipients: any): void
        getCCRecipients(): any
        setCCRecipients(cCRecipients: any): void
        getEnclosures(): any
        setEnclosures(enclosures: any): void
        isReceived(): boolean
        getRecipients(): any
        setRecipients(recipients: any): void
        getSendDateTime(): any
        getSender(): string
        getSubject(): string
        setSubject(subject: string): void
    }
    export interface Mailer {
        getBCCRecipients(): any
        setBCCRecipients(bCCRecipients: any): void
        getCCRecipients(): any
        setCCRecipients(cCRecipients: any): void
        getEnclosures(): any
        setEnclosures(enclosures: any): void
        isReceived(): boolean
        getRecipients(): any
        setRecipients(recipients: any): void
        getSendDateTime(): any
        getSender(): string
        getSubject(): string
        setSubject(subject: string): void
    }
    export interface MailingLabel {
        printOut(name: any, address: any, extractAddress: any, laserTray: any, singleLabel: any, row: any, column: any, printEPostageLabel: any, vertical: any): void
        createNewDocument(name: any, address: any, autoText: any, extractAddress: any, laserTray: any, printEPostageLabel: any, vertical: any): word.Document
        setDefaultPrintBarCode(defaultPrintBarCode: boolean): void
        isDefaultPrintBarCode(): boolean
        setVertical(vertical: boolean): void
        isVertical(): boolean
        getDefaultLabelName(): string
        setDefaultLabelName(defaultLabelName: string): void
        getDefaultLaserTray(): int
        setDefaultLaserTray(defaultLaserTray: int): void
        getCustomLabels(): word.CustomLabels
        labelOptions(): void
    }
    export interface MailingLabel {
        printOut(name: any, address: any, extractAddress: any, laserTray: any, singleLabel: any, row: any, column: any, printEPostageLabel: any, vertical: any): void
        createNewDocument(name: any, address: any, autoText: any, extractAddress: any, laserTray: any, printEPostageLabel: any, vertical: any): word.Document
        setDefaultPrintBarCode(defaultPrintBarCode: boolean): void
        isDefaultPrintBarCode(): boolean
        setVertical(vertical: boolean): void
        isVertical(): boolean
        getDefaultLabelName(): string
        setDefaultLabelName(defaultLabelName: string): void
        getDefaultLaserTray(): int
        setDefaultLaserTray(defaultLaserTray: int): void
        getCustomLabels(): word.CustomLabels
        labelOptions(): void
    }
    export interface MailMerge {
        getFields(): word.MailMergeFields
        execute(pause: any): void
        check(): void
        getState(): int
        isViewMailMergeFieldCodes(): boolean
        setViewMailMergeFieldCodes(viewMailMergeFieldCodes: int): void
        isHighlightMergeFields(): boolean
        setHighlightMergeFields(highlightMergeFields: boolean): void
        getMailAddressFieldName(): string
        setMailAddressFieldName(mailAddressFieldName: string): void
        isMailAsAttachment(): boolean
        setMailAsAttachment(mailAsAttachment: boolean): void
        getMainDocumentType(): int
        setMainDocumentType(mainDocumentType: int): void
        getShowSendToCustom(): string
        setShowSendToCustom(showSendToCustom: string): void
        isSuppressBlankLines(): boolean
        setSuppressBlankLines(suppressBlankLines: boolean): void
        createHeaderSource(name: string, passwordDocument: any, writePasswordDocument: any, headerRecord: any): void
        getDataSource(): word.MailMergeDataSource
        getDestination(): int
        setDestination(destination: int): void
        getMailFormat(): int
        setMailFormat(mailFormat: int): void
        getMailSubject(): string
        setMailSubject(mailSubject: string): void
        getWizardState(): int
        setWizardState(wizardState: int): void
        createDataSource(name: any, passwordDocument: any, writePasswordDocument: any, headerRecord: any, mSQuery: any, sqlStatement: any, sqlStatement1: any, connection: any, linkToSource: any): void
        editDataSource(): void
        editHeaderSource(): void
        editMainDocument(): void
        openDataSource(name: string, format: any, confirmConversions: any, readOnly: any, linkToSource: any, addToRecentFiles: any, passwordDocument: any, passwordTemplate: any, revert: any, writePasswordDocument: any, writePasswordTemplate: any, connection: any, sqlStatement: any, sqlStatement1: any, openExclusive: any, subType: any): void
        openHeaderSource(name: string, format: int, confirmConversions: boolean, readOnly: boolean, addToRecentFiles: boolean, passwordDocument: string, passwordTemplate: string, revert: boolean, writePasswordDocument: string, writePasswordTemplate: string, openExclusive: boolean): void
        showWizard(initialState: any, showDocumentStep: any, showTemplateStep: any, showDataStep: any, showWriteStep: any, showPreviewStep: any, showMergeStep: any): void
    }
    export interface MailMerge {
        getFields(): word.MailMergeFields
        execute(pause: any): void
        check(): void
        getState(): int
        isViewMailMergeFieldCodes(): boolean
        setViewMailMergeFieldCodes(viewMailMergeFieldCodes: int): void
        isHighlightMergeFields(): boolean
        setHighlightMergeFields(highlightMergeFields: boolean): void
        getMailAddressFieldName(): string
        setMailAddressFieldName(mailAddressFieldName: string): void
        isMailAsAttachment(): boolean
        setMailAsAttachment(mailAsAttachment: boolean): void
        getMainDocumentType(): int
        setMainDocumentType(mainDocumentType: int): void
        getShowSendToCustom(): string
        setShowSendToCustom(showSendToCustom: string): void
        isSuppressBlankLines(): boolean
        setSuppressBlankLines(suppressBlankLines: boolean): void
        createHeaderSource(name: string, passwordDocument: any, writePasswordDocument: any, headerRecord: any): void
        getDataSource(): word.MailMergeDataSource
        getDestination(): int
        setDestination(destination: int): void
        getMailFormat(): int
        setMailFormat(mailFormat: int): void
        getMailSubject(): string
        setMailSubject(mailSubject: string): void
        getWizardState(): int
        setWizardState(wizardState: int): void
        createDataSource(name: any, passwordDocument: any, writePasswordDocument: any, headerRecord: any, mSQuery: any, sqlStatement: any, sqlStatement1: any, connection: any, linkToSource: any): void
        editDataSource(): void
        editHeaderSource(): void
        editMainDocument(): void
        openDataSource(name: string, format: any, confirmConversions: any, readOnly: any, linkToSource: any, addToRecentFiles: any, passwordDocument: any, passwordTemplate: any, revert: any, writePasswordDocument: any, writePasswordTemplate: any, connection: any, sqlStatement: any, sqlStatement1: any, openExclusive: any, subType: any): void
        openHeaderSource(name: string, format: int, confirmConversions: boolean, readOnly: boolean, addToRecentFiles: boolean, passwordDocument: string, passwordTemplate: string, revert: boolean, writePasswordDocument: string, writePasswordTemplate: string, openExclusive: boolean): void
        showWizard(initialState: any, showDocumentStep: any, showTemplateStep: any, showDataStep: any, showWriteStep: any, showPreviewStep: any, showMergeStep: any): void
    }
    export interface MailMergeDataField {
        getName(): string
        getValue(): string
        getIndex(): int
    }
    export interface MailMergeDataField {
        getName(): string
        getValue(): string
        getIndex(): int
    }
    export interface MailMergeDataFields {
        item(index: int): word.MailMergeDataField
        getCount(): int
    }
    export interface MailMergeDataFields {
        item(index: int): word.MailMergeDataField
        getCount(): int
    }
    export interface MailMergeDataSource {
        getName(): string
        close(): void
        getType(): int
        getTableName(): string
        getHeaderSourceName(): string
        getHeaderSourceType(): int
        setInvalidAddress(invalidAddress: boolean): void
        getInvalidComments(): string
        setInvalidComments(invalidComments: string): void
        getMappedDataFields(): word.MappedDataFields
        setAllIncludedFlags(included: boolean): void
        getActiveRecord(): int
        setActiveRecord(activeRecord: int): void
        getConnectString(): string
        getDataFields(): word.MailMergeDataFields
        getFieldNames(): word.MailMergeFieldNames
        getFirstRecord(): int
        setFirstRecord(firstRecord: int): void
        isIncluded(): boolean
        setIncluded(included: boolean): void
        isInvalidAddress(): boolean
        getLastRecord(): int
        setLastRecord(lastRecord: int): void
        getQueryString(): string
        setQueryString(queryString: string): void
        getRecordCount(): int
        findRecord(findText: string, field: string): boolean
        setAllErrorFlags(invalid: boolean, invalidComment: string): void
    }
    export interface MailMergeDataSource {
        getName(): string
        close(): void
        getType(): int
        getTableName(): string
        getHeaderSourceName(): string
        getHeaderSourceType(): int
        setInvalidAddress(invalidAddress: boolean): void
        getInvalidComments(): string
        setInvalidComments(invalidComments: string): void
        getMappedDataFields(): word.MappedDataFields
        setAllIncludedFlags(included: boolean): void
        getActiveRecord(): int
        setActiveRecord(activeRecord: int): void
        getConnectString(): string
        getDataFields(): word.MailMergeDataFields
        getFieldNames(): word.MailMergeFieldNames
        getFirstRecord(): int
        setFirstRecord(firstRecord: int): void
        isIncluded(): boolean
        setIncluded(included: boolean): void
        isInvalidAddress(): boolean
        getLastRecord(): int
        setLastRecord(lastRecord: int): void
        getQueryString(): string
        setQueryString(queryString: string): void
        getRecordCount(): int
        findRecord(findText: string, field: string): boolean
        setAllErrorFlags(invalid: boolean, invalidComment: string): void
    }
    export interface MailMergeField {
        delete(): void
        copy(): void
        getType(): int
        getResult(): word.Range
        isLocked(): boolean
        cut(): void
        getPrevious(): word.MailMergeField
        select(): void
        getNext(): word.MailMergeField
        getCode(): word.Range
        setCode(code: word.Range): void
        setLocked(locked: boolean): void
    }
    export interface MailMergeField {
        delete(): void
        copy(): void
        getType(): int
        getResult(): word.Range
        isLocked(): boolean
        cut(): void
        getPrevious(): word.MailMergeField
        select(): void
        getNext(): word.MailMergeField
        getCode(): word.Range
        setCode(code: word.Range): void
        setLocked(locked: boolean): void
    }
    export interface MailMergeFieldName {
        getName(): string
        getIndex(): int
    }
    export interface MailMergeFieldName {
        getName(): string
        getIndex(): int
    }
    export interface MailMergeFieldNames {
        item(index: any): word.MailMergeFieldName
        getCount(): int
    }
    export interface MailMergeFieldNames {
        item(index: any): word.MailMergeFieldName
        getCount(): int
    }
    export interface MailMergeFields {
        add(range: word.Range, name: string): word.MailMergeField
        item(index: int): word.MailMergeField
        getCount(): int
        addAsk(range: word.Range, name: string, prompt: any, defaultAskText: any, askOnce: any): word.MailMergeField
        addIf(range: word.Range, mergeField: string, comparison: int, compareTo: any, trueAutoText: any, trueText: any, falseAutoText: any, falseText: any): word.MailMergeField
        addNext(range: word.Range): word.MailMergeField
        addSet(range: word.Range, name: string, valueText: any, valueAutoText: any): word.MailMergeField
        addFillIn(range: word.Range, prompt: any, defaultFillInText: any, askOnce: any): word.MailMergeField
        addMergeRec(range: word.Range): word.MailMergeField
        addMergeSeq(range: word.Range): word.MailMergeField
        addNextIf(range: word.Range, MergeField: string, comparison: int, compareTo: any): word.MailMergeField
        addSkipIf(range: word.Range, mergeField: string, comparison: int, compareTo: any): word.MailMergeField
    }
    export interface MailMergeFields {
        add(range: word.Range, name: string): word.MailMergeField
        item(index: int): word.MailMergeField
        getCount(): int
        addAsk(range: word.Range, name: string, prompt: any, defaultAskText: any, askOnce: any): word.MailMergeField
        addIf(range: word.Range, mergeField: string, comparison: int, compareTo: any, trueAutoText: any, trueText: any, falseAutoText: any, falseText: any): word.MailMergeField
        addNext(range: word.Range): word.MailMergeField
        addSet(range: word.Range, name: string, valueText: any, valueAutoText: any): word.MailMergeField
        addFillIn(range: word.Range, prompt: any, defaultFillInText: any, askOnce: any): word.MailMergeField
        addMergeRec(range: word.Range): word.MailMergeField
        addMergeSeq(range: word.Range): word.MailMergeField
        addNextIf(range: word.Range, MergeField: string, comparison: int, compareTo: any): word.MailMergeField
        addSkipIf(range: word.Range, mergeField: string, comparison: int, compareTo: any): word.MailMergeField
    }
    export interface MailMessage {
        checkName(): void
        delete(): void
        forward(): void
        goToNext(): void
        reply(): void
        replyAll(): void
        displayMoveDialog(): void
        displayProperties(): void
        displaySelectNamesDialog(): void
        goToPrevious(): void
        toggleHeader(): void
    }
    export interface MailMessage {
        checkName(): void
        delete(): void
        forward(): void
        goToNext(): void
        reply(): void
        replyAll(): void
        displayMoveDialog(): void
        displayProperties(): void
        displaySelectNamesDialog(): void
        goToPrevious(): void
        toggleHeader(): void
    }
    export interface MappedDataField {
        getName(): string
        getValue(): string
        getIndex(): int
        getDataFieldIndex(): int
        setDataFieldIndex(dataFieldIndex: int): void
        getDataFieldName(): string
        setDataFieldName(dataFieldName: string): void
    }
    export interface MappedDataField {
        getName(): string
        getValue(): string
        getIndex(): int
        getDataFieldIndex(): int
        setDataFieldIndex(dataFieldIndex: int): void
        getDataFieldName(): string
        setDataFieldName(dataFieldName: string): void
    }
    export interface MappedDataFields {
        item(index: int): word.MappedDataField
        getCount(): int
    }
    export interface MappedDataFields {
        item(index: int): word.MappedDataField
        getCount(): int
    }
    export interface Note {
        delete(): void
        getIndex(): int
        getRange(): word.Range
        getReference(): word.Range
        getMNote(): any
    }
    export interface Note {
        delete(): void
        getIndex(): int
        getRange(): word.Range
        getReference(): word.Range
        getMNote(): any
    }
    export interface Notes {
        add(range: word.Range, Reference: string, text: string): word.Note
        getLocation(): int
        getSeparator(): word.Range
        convert(): void
        item(index: int): word.Note
        getCount(): int
        setLocation(location: int): void
        setNumberStyle(style: int): void
        getNumberStyle(): int
        resetContinuationSeparator(): void
        swapWithEndnotesAndFootnotes(): void
        getNotes(): any
        getContinuationNotice(): word.Range
        setStartingNumber(number: int): void
        getStartingNumber(): int
        getContinuationSeparator(): word.Range
        resetContinuationNotice(): void
        setNumberingRule(rule: int): void
        getNumberingRule(): int
        resetSeparator(): void
    }
    export interface Notes {
        add(range: word.Range, Reference: string, text: string): word.Note
        getLocation(): int
        getSeparator(): word.Range
        convert(): void
        item(index: int): word.Note
        getCount(): int
        setLocation(location: int): void
        setNumberStyle(style: int): void
        getNumberStyle(): int
        resetContinuationSeparator(): void
        swapWithEndnotesAndFootnotes(): void
        getNotes(): any
        getContinuationNotice(): word.Range
        setStartingNumber(number: int): void
        getStartingNumber(): int
        getContinuationSeparator(): word.Range
        resetContinuationNotice(): void
        setNumberingRule(rule: int): void
        getNumberingRule(): int
        resetSeparator(): void
    }
    export interface OLEControl {
        getName(): string
        delete(): void
        setName(name: string): void
        copy(): void
        getWidth(): double
        getLeft(): double
        setLeft(left: double): void
        setTop(top: double): void
        getTop(): double
        setWidth(width: double): void
        activate(): void
        getHeight(): float
        setHeight(height: double): void
        cut(): void
        select(): void
        doAction(type: int): void
        getIDispatch(): long
        getAutomation(): any
    }
    export interface OLEControl {
        getName(): string
        delete(): void
        setName(name: string): void
        copy(): void
        getWidth(): double
        getLeft(): double
        setLeft(left: double): void
        setTop(top: double): void
        getTop(): double
        setWidth(width: double): void
        activate(): void
        getHeight(): float
        setHeight(height: double): void
        cut(): void
        select(): void
        doAction(type: int): void
        getIDispatch(): long
        getAutomation(): any
    }
    export interface OLEFormat {
        getObject(): any
        open(): void
        activate(): void
        edit(): void
        getLabel(): string
        doVerb(verbIndex: any): void
        getClassType(): string
        setClassType(classType: string): void
        isDisplayAsIcon(): boolean
        setDisplayAsIcon(displayAsIcon: boolean): void
        getIconIndex(): int
        setIconIndex(iconIndex: int): void
        getIconLabel(): string
        setIconLabel(iconLabel: string): void
        getIconName(): string
        setIconName(iconName: string): void
        getIconPath(): string
        getProgID(): string
        activateAs(classType: string): void
        convertTo(classType: any, displayAsIcon: any, iconFileName: any, iconIndex: any, iconLabel: any): void
        isPreserveFormattingOnUpdate(): boolean
        setPreserveFormattingOnUpdate(preserveFormattingOnUpdate: boolean): void
    }
    export interface OLEFormat {
        getObject(): any
        open(): void
        activate(): void
        edit(): void
        getLabel(): string
        doVerb(verbIndex: any): void
        getClassType(): string
        setClassType(classType: string): void
        isDisplayAsIcon(): boolean
        setDisplayAsIcon(displayAsIcon: boolean): void
        getIconIndex(): int
        setIconIndex(iconIndex: int): void
        getIconLabel(): string
        setIconLabel(iconLabel: string): void
        getIconName(): string
        setIconName(iconName: string): void
        getIconPath(): string
        getProgID(): string
        activateAs(classType: string): void
        convertTo(classType: any, displayAsIcon: any, iconFileName: any, iconIndex: any, iconLabel: any): void
        isPreserveFormattingOnUpdate(): boolean
        setPreserveFormattingOnUpdate(preserveFormattingOnUpdate: boolean): void
    }
    export interface Options {
        setColor(ywColorIndex: int, what: int, setting: number[]): number[]
        getDefaultFilePath(path: int): string
        setShowErrorDialog(showFlag: boolean): boolean
        isShowErrorDialog(): boolean
        isAddBiDirectionalMarksWhenSavingTextFile(): boolean
        setAddBiDirectionalMarksWhenSavingTextFile(addBiDirectionalMarksWhenSavingTextFile: boolean): void
        isAutoFormatAsYouTypeFormatListItemBeginning(): boolean
        setAutoFormatAsYouTypeFormatListItemBeginning(autoFormatAsYouTypeFormatListItemBeginning: boolean): void
        isAutoFormatAsYouTypeReplaceFarEastDashes(): boolean
        setAutoFormatAsYouTypeReplaceFarEastDashes(autoFormatAsYouTypeReplaceFarEastDashes: boolean): void
        setLastEditPaneEnabled(enable: boolean): void
        getDeletedTextColor(): int
        setDeletedTextColor(deletedTextColor: int): void
        getDeletedTextMark(): int
        setDeletedTextMark(deletedTextMark: int): void
        getDiacriticColorVal(): int
        setDiacriticColorVal(diacriticColorVal: int): void
        isDisplayGridLines(): boolean
        setDisplayGridLines(displayGridLines: boolean): void
        isDisplayPasteOptions(): boolean
        setDisplayPasteOptions(displayPasteOptions: boolean): void
        isDisplaySmartTagButtons(): boolean
        getDocumentViewDirection(): int
        setDocumentViewDirection(documentViewDirection: int): void
        setEormatScanning(eormatScanning: boolean): void
        getGridDistanceVertical(): double
        setGridDistanceVertical(gridDistanceVertical: double): void
        getGridOriginHorizontal(): double
        setGridOriginHorizontal(gridOriginHorizontal: double): void
        getGridOriginVertical(): double
        setGridOriginVertical(gridOriginVertical: double): void
        isIgnoreMixedDigits(): boolean
        setIgnoreMixedDigits(ignoreMixedDigits: boolean): void
        isIgnoreUppercase(): boolean
        setIgnoreUppercase(ignoreUppercase: boolean): void
        isIMEAutomaticControl(): boolean
        setIMEAutomaticControl(automaticControl: boolean): void
        isInlineConversion(): boolean
        setInlineConversion(inlineConversion: boolean): void
        getInsertedTextColor(): int
        setInsertedTextColor(insertedTextColor: int): void
        getInsertedTextMark(): int
        setInsertedTextMark(insertedTextMark: int): void
        setINSKeyForPaste(keyForPaste: boolean): void
        getInterpretHighAnsi(): int
        setInterpretHighAnsi(interpretHighAnsi: int): void
        setLabelSmartTags(labelSmartTags: boolean): void
        isLocalNetworkFile(): boolean
        setLocalNetworkFile(localNetworkFile: boolean): void
        setMatchFuzzyByte(matchFuzzyByte: boolean): void
        isDisableFeaturesbyDefault(): boolean
        setDisableFeaturesbyDefault(disableFeaturesbyDefault: boolean): void
        setDisplaySmartTagButtons(displaySmartTagButtons: boolean): void
        isEnableMisusedWordsDictionary(): boolean
        setEnableMisusedWordsDictionary(enableMisusedWordsDictionary: boolean): void
        isEnvelopeFeederInstalled(): boolean
        setEnvelopeFeederInstalled(envelopeFeederInstalled: boolean): void
        getGridDistanceHorizontal(): double
        setGridDistanceHorizontal(gridDistanceHorizontal: double): void
        isHangulHanjaFastConversion(): boolean
        setHangulHanjaFastConversion(hangulHanjaFastConversion: boolean): void
        isIgnoreInternetAndFileAddresses(): boolean
        isMatchFuzzyIterationMark(): boolean
        setMatchFuzzyIterationMark(matchFuzzyIterationMark: boolean): void
        isEnableHangulHanjaRecentOrdering(): boolean
        setEnableHangulHanjaRecentOrdering(enableHangulHanjaRecentOrdering: boolean): void
        setIgnoreInternetAndFileAddresses(ignoreInternetAndFileAddresses: boolean): void
        setMatchFuzzyCase(matchFuzzyCase: boolean): void
        setMatchFuzzyDash(matchFuzzyDash: boolean): void
        isMatchFuzzyHiragana(): boolean
        setMatchFuzzyHiragana(matchFuzzyHiragana: boolean): void
        isMatchFuzzyKanji(): boolean
        setMatchFuzzyKanji(matchFuzzyKanji: boolean): void
        setMatchFuzzyKiKu(matchFuzzyKiKu: boolean): void
        isMatchFuzzyOldKana(): boolean
        setMatchFuzzyOldKana(matchFuzzyOldKana: boolean): void
        isMatchFuzzyPunctuation(): boolean
        setMatchFuzzyPunctuation(matchFuzzyPunctuation: boolean): void
        isMatchFuzzySmallKana(): boolean
        setMatchFuzzySmallKana(matchFuzzySmallKana: boolean): void
        isMatchFuzzySpace(): boolean
        setMatchFuzzySpace(matchFuzzySpace: boolean): void
        getMeasurementUnit(): int
        setMeasurementUnit(measurementUnit: int): void
        isPasteAdjustWordSpacing(): boolean
        isPasteMergeFromPPT(): boolean
        setPasteMergeFromPPT(pasteMergeFromPPT: boolean): void
        isPasteMergeFromXL(): boolean
        setPasteMergeFromXL(pasteMergeFromXL: boolean): void
        isPasteMergeLists(): boolean
        setPasteMergeLists(pasteMergeLists: boolean): void
        isPasteSmartCutPaste(): boolean
        setPasteSmartCutPaste(pasteSmartCutPaste: boolean): void
        getPictureWrapType(): int
        setPictureWrapType(pictureWrapType: int): void
        isPrintBackground(): boolean
        setPrintBackground(printBackground: boolean): void
        isPrintBackgrounds(): boolean
        setPrintBackgrounds(printBackgrounds: boolean): void
        setPrintHiddenText(printHiddenText: boolean): void
        isPrintDrawingObjects(): boolean
        setPrintDrawingObjects(printDrawingObjects: boolean): void
        isPrintFieldCodes(): boolean
        setPrintFieldCodes(printFieldCodes: boolean): void
        isPrintHiddenText(): boolean
        isMatchFuzzyProintedSoundMark(): boolean
        setMatchFuzzyProintedSoundMark(matchFuzzyProintedSoundMark: boolean): void
        getMultipleWordConversionsMode(): int
        setMultipleWordConversionsMode(multipleWordConversionsMode: int): void
        isOptimizeForWord97byDefault(): boolean
        setOptimizeForWord97byDefault(optimizeForWord97byDefault: boolean): void
        isPasteAdjustParagraphSpacing(): boolean
        setPasteAdjustParagraphSpacing(pasteAdjustParagraphSpacing: boolean): void
        isPasteAdjustTableFormatting(): boolean
        setPasteAdjustTableFormatting(pasteAdjustTableFormatting: boolean): void
        setPasteAdjustWordSpacing(pasteAdjustWordSpacing: boolean): void
        isPasteSmartStyleBehavior(): boolean
        setPasteSmartStyleBehavior(pasteSmartStyleBehavior: boolean): void
        isPrintEvenPagesInAscendingOrder(): boolean
        isPrintOddPagesInAscendingOrder(): boolean
        setPrintOddPagesInAscendingOrder(printOddPagesInAscendingOrder: boolean): void
        getRevisedPropertiesColor(): int
        setRevisedPropertiesColor(revisedPropertiesColor: int): void
        setPrintEvenPagesInAscendingOrder(printEvenPagesInAscendingOrder: boolean): void
        getRevisionsBalloonPrintOrientation(): int
        setRevisionsBalloonPrintOrientation(revisionsBalloonPrintOrientation: int): void
        getColor(what: int): int
        isWPHelp(): boolean
        setAllowAccentedUppercase(allowAccentedUppercase: boolean): void
        isAllowFastSave(): boolean
        setAllowFastSave(allowFastSave: boolean): void
        getArabicMode(): int
        setArabicMode(arabicMode: int): void
        getArabicNumeral(): int
        setArabicNumeral(arabicNumeral: int): void
        isBackgroundOpen(): boolean
        isBackgroundSave(): boolean
        isBlueScreen(): boolean
        setBlueScreen(blueScreen: boolean): void
        getCommentsColor(): int
        setCommentsColor(commentsColor: int): void
        isCreateBackup(): boolean
        setCreateBackup(createBackup: boolean): void
        getDefaultTray(): string
        setDefaultTray(defaultTray: string): void
        getDefaultTrayID(): int
        setDefaultTrayID(defaultTrayID: int): void
        isEnableSound(): boolean
        setEnableSound(enableSound: boolean): void
        isEormatScanning(): boolean
        getHebrewMode(): int
        setHebrewMode(hebrewMode: int): void
        isINSKeyForPaste(): boolean
        isLabelSmartTags(): boolean
        isMapPaperSize(): boolean
        setMapPaperSize(mapPaperSize: boolean): void
        isMatchFuzzyAY(): boolean
        setMatchFuzzyAY(matchFuzzyAY: boolean): void
        isMatchFuzzyBV(): boolean
        setMatchFuzzyBV(matchFuzzyBV: boolean): void
        isMatchFuzzyByte(): boolean
        isMatchFuzzyCase(): boolean
        isMatchFuzzyDash(): boolean
        isMatchFuzzyDZ(): boolean
        setMatchFuzzyDZ(matchFuzzyDZ: boolean): void
        isMatchFuzzyHF(): boolean
        setMatchFuzzyHF(matchFuzzyHF: boolean): void
        isMatchFuzzyKiKu(): boolean
        isMatchFuzzyTC(): boolean
        setMatchFuzzyTC(matchFuzzyTC: boolean): void
        isMatchFuzzyZJ(): boolean
        setMatchFuzzyZJ(matchFuzzyZJ: boolean): void
        getMonthNames(): int
        setMonthNames(monthNames: int): void
        isOvertype(): boolean
        setOvertype(overtype: boolean): void
        isPagination(): boolean
        setPagination(pagination: boolean): void
        getPictureEditor(): string
        setPictureEditor(pictureEditor: string): void
        isPrintComments(): boolean
        setPrintComments(printComments: boolean): void
        isPrintDraft(): boolean
        setPrintDraft(printDraft: boolean): void
        isPrintReverse(): boolean
        setPrintReverse(printReverse: boolean): void
        isAutoFormatAsYouTypeApplyBorders(): boolean
        setAutoFormatAsYouTypeApplyBorders(autoFormatAsYouTypeApplyBorders: boolean): void
        isAutoFormatAsYouTypeApplyBulletedLists(): boolean
        setAutoFormatAsYouTypeApplyBulletedLists(autoFormatAsYouTypeApplyBulletedLists: boolean): void
        isAutoFormatAsYouTypeApplyClosings(): boolean
        setAutoFormatAsYouTypeApplyClosings(autoFormatAsYouTypeApplyClosings: boolean): void
        isAutoFormatAsYouTypeApplyFirstIndents(): boolean
        setAutoFormatAsYouTypeApplyFirstIndents(autoFormatAsYouTypeApplyFirstIndents: boolean): void
        isAutoFormatAsYouTypeApplyHeadings(): boolean
        setAutoFormatAsYouTypeApplyHeadings(autoFormatAsYouTypeApplyHeadings: boolean): void
        isAutoFormatAsYouTypeApplyNumberedLists(): boolean
        setAutoFormatAsYouTypeApplyNumberedLists(autoFormatAsYouTypeApplyNumberedLists: boolean): void
        setAutoFormatAsYouTypeApplyTables(autoFormatAsYouTypeApplyTables: boolean): void
        isAutoFormatAsYouTypeAutoLetterWizard(): boolean
        setAutoFormatAsYouTypeAutoLetterWizard(autoFormatAsYouTypeAutoLetterWizard: boolean): void
        isAutoFormatAsYouTypeDefineStyles(): boolean
        setAutoFormatAsYouTypeDefineStyles(autoFormatAsYouTypeDefineStyles: boolean): void
        isAutoFormatAsYouTypeDeleteAutoSpaces(): boolean
        setAutoFormatAsYouTypeDeleteAutoSpaces(autoFormatAsYouTypeDeleteAutoSpaces: boolean): void
        isAutoFormatAsYouTypeInsertClosings(): boolean
        setAutoFormatAsYouTypeInsertClosings(autoFormatAsYouTypeInsertClosings: boolean): void
        setAutoFormatAsYouTypeInsertOvers(autoFormatAsYouTypeInsertOvers: boolean): void
        isAutoFormatAsYouTypeMatchParentheses(): boolean
        isAddControlCharacters(): boolean
        setAddControlCharacters(addControlCharacters: boolean): void
        isAddHebDoubleQuote(): boolean
        setAddHebDoubleQuote(addHebDoubleQuote: boolean): void
        isAllowAccentedUppercase(): boolean
        isAllowClickAndTypeMouse(): boolean
        isAllowDragAndDrop(): boolean
        setAllowDragAndDrop(allowDragAndDrop: boolean): void
        isAllowPixelUnits(): boolean
        setAllowPixelUnits(allowPixelUnits: boolean): void
        setAllowClickAndTypeMouse(allowClickAndTypeMouse: boolean): void
        isAllowCombinedAuxiliaryForms(): boolean
        setAllowCombinedAuxiliaryForms(allowCombinedAuxiliaryForms: boolean): void
        isAllowCompoundNounProcessing(): boolean
        setAllowCompoundNounProcessing(allowCompoundNounProcessing: boolean): void
        setAnimateScreenMovements(animateScreenMovements: boolean): void
        isApplyFarEastFontsToAscii(): boolean
        setApplyFarEastFontsToAscii(applyFarEastFontsToAscii: boolean): void
        isAutoFormatApplyBulletedLists(): boolean
        setAutoFormatApplyBulletedLists(autoFormatApplyBulletedLists: boolean): void
        isAutoFormatApplyFirstIndents(): boolean
        setAutoFormatApplyFirstIndents(autoFormatApplyFirstIndents: boolean): void
        isAutoFormatApplyHeadings(): boolean
        setAutoFormatApplyHeadings(autoFormatApplyHeadings: boolean): void
        isAutoFormatApplyOtherParas(): boolean
        setAutoFormatApplyOtherParas(autoFormatApplyOtherParas: boolean): void
        isAutoFormatAsYouTypeApplyDates(): boolean
        setAutoFormatAsYouTypeApplyDates(autoFormatAsYouTypeApplyDates: boolean): void
        isAutoFormatAsYouTypeApplyTables(): boolean
        isAutoFormatAsYouTypeInsertOvers(): boolean
        isAutoFormatDeleteAutoSpaces(): boolean
        setAutoFormatDeleteAutoSpaces(autoFormatDeleteAutoSpaces: boolean): void
        isAutoFormatMatchParentheses(): boolean
        setAutoFormatMatchParentheses(autoFormatMatchParentheses: boolean): void
        isAutoFormatPlainTextWordMail(): boolean
        setAutoFormatPlainTextWordMail(autoFormatPlainTextWordMail: boolean): void
        isAutoFormatPreserveStyles(): boolean
        setAutoFormatPreserveStyles(autoFormatPreserveStyles: boolean): void
        isAutoFormatReplaceFarEastDashes(): boolean
        isAutoFormatReplaceFractions(): boolean
        setAutoFormatReplaceFractions(autoFormatReplaceFractions: boolean): void
        isAutoFormatReplaceHyperlinks(): boolean
        setAutoFormatReplaceHyperlinks(autoFormatReplaceHyperlinks: boolean): void
        isAutoFormatReplaceOrdinals(): boolean
        setAutoFormatReplaceOrdinals(autoFormatReplaceOrdinals: boolean): void
        isAutoFormatReplaceQuotes(): boolean
        setAutoFormatReplaceQuotes(autoFormatReplaceQuotes: boolean): void
        isAllowReadingMode(): boolean
        setAllowReadingMode(allowReadingMode: boolean): void
        isAnimateScreenMovements(): boolean
        isAutoCreateNewDrawings(): boolean
        setAutoCreateNewDrawings(autoCreateNewDrawings: boolean): void
        isAutoFormatApplyLists(): boolean
        setAutoFormatApplyLists(autoFormatApplyLists: boolean): void
        isAutoKeyboardSwitching(): boolean
        setAutoKeyboardSwitching(autoKeyboardSwitching: boolean): void
        isAutoWordSelection(): boolean
        setAutoWordSelection(autoWordSelection: boolean): void
        setBackgroundOpen(backgroundOpen: boolean): void
        setBackgroundSave(backgroundSave: boolean): void
        getButtonFieldClicks(): int
        setButtonFieldClicks(buttonFieldClicks: int): void
        isCheckGrammarAsYouType(): boolean
        setCheckGrammarAsYouType(checkGrammarAsYouType: boolean): void
        isCheckHangulEndings(): boolean
        setCheckHangulEndings(checkHangulEndings: boolean): void
        isCheckSpellingAsYouType(): boolean
        applyTrackChangeSetting(setting: number[]): void
        isConfirmConversions(): boolean
        setConfirmConversions(confirmConversions: boolean): void
        getCursorMovement(): int
        setCursorMovement(cursorMovement: int): void
        getDefaultBorderColor(): int
        setDefaultBorderColor(defaultBorderColor: int): void
        getDefaultEPostageApp(): string
        setDefaultEPostageApp(defaultEPostageApp: string): void
        setDefaultFilePath(path: int, defaultFilePath: string): void
        getDefaultOpenFormat(): int
        setDefaultOpenFormat(defaultOpenFormat: int): void
        getDefaultTextEncoding(): int
        setDefaultTextEncoding(defaultTextEncoding: int): void
        setAutoFormatAsYouTypeMatchParentheses(autoFormatAsYouTypeMatchParentheses: boolean): void
        isAutoFormatAsYouTypeReplaceFractions(): boolean
        setAutoFormatAsYouTypeReplaceFractions(autoFormatAsYouTypeReplaceFractions: boolean): void
        isAutoFormatAsYouTypeReplaceHyperlinks(): boolean
        setAutoFormatAsYouTypeReplaceHyperlinks(autoFormatAsYouTypeReplaceHyperlinks: boolean): void
        isAutoFormatAsYouTypeReplaceOrdinals(): boolean
        setAutoFormatAsYouTypeReplaceOrdinals(autoFormatAsYouTypeReplaceOrdinals: boolean): void
        isAutoFormatAsYouTypeReplaceQuotes(): boolean
        setAutoFormatAsYouTypeReplaceQuotes(autoFormatAsYouTypeReplaceQuotes: boolean): void
        isAutoFormatAsYouTypeReplaceSymbols(): boolean
        setAutoFormatAsYouTypeReplaceSymbols(autoFormatAsYouTypeReplaceSymbols: boolean): void
        setAutoFormatReplaceFarEastDashes(autoFormatReplaceFarEastDashes: boolean): void
        isAutoFormatAsYouTypeReplacePlainTextEmphasis(): boolean
        setAutoFormatAsYouTypeReplacePlainTextEmphasis(autoFormatAsYouTypeReplacePlainTextEmphasis: boolean): void
        getDisableFeaturesIntroducedAfterbyDefault(): int
        setDisableFeaturesIntroducedAfterbyDefault(disableFeaturesIntroducedAfterbyDefault: int): void
        isAutoFormatReplacePlainTextEmphasis(): boolean
        setAutoFormatReplacePlainTextEmphasis(autoFormatReplacePlainTextEmphasis: boolean): void
        isAutoFormatReplaceSymbols(): boolean
        setAutoFormatReplaceSymbols(autoFormatReplaceSymbols: boolean): void
        isCheckGrammarWithSpelling(): boolean
        setCheckGrammarWithSpelling(checkGrammarWithSpelling: boolean): void
        setCheckSpellingAsYouType(checkSpellingAsYouType: boolean): void
        getapplyTrackChangeSetting(): number[]
        isConvertHighAnsiToFarEast(): boolean
        setConvertHighAnsiToFarEast(convertHighAnsiToFarEast: boolean): void
        isCtrlClickHyperlinkToOpen(): boolean
        setCtrlClickHyperlinkToOpen(ctrlClickHyperlinkToOpen: boolean): void
        getDefaultBorderColorIndex(): int
        setDefaultBorderColorIndex(defaultBorderColorIndex: int): void
        getDefaultBorderLineStyle(): int
        setDefaultBorderLineStyle(defaultBorderLineStyle: int): void
        getDefaultBorderLineWidth(): int
        setDefaultBorderLineWidth(defaultBorderLineWidth: int): void
        getDefaultHighlightColorIndex(): int
        setDefaultHighlightColorIndex(defaultHighlightColorIndex: int): void
        isPrintProperties(): boolean
        setPrintProperties(printProperties: boolean): void
        isPromptUpdateStyle(): boolean
        setPromptUpdateStyle(promptUpdateStyle: boolean): void
        isReplaceSelection(): boolean
        setReplaceSelection(replaceSelection: boolean): void
        getRevisedLinesColor(): int
        setRevisedLinesColor(revisedLinesColor: int): void
        getRevisedLinesMark(): int
        setRevisedLinesMark(revisedLinesMark: int): void
        getRevisedPropertiesMark(): int
        setRevisedPropertiesMark(revisedPropertiesMark: int): void
        setRTFInClipboard(inClipboard: boolean): void
        isSaveNormalPrompt(): boolean
        setSaveNormalPrompt(saveNormalPrompt: boolean): void
        isSavePropertiesPrompt(): boolean
        setSavePropertiesPrompt(savePropertiesPrompt: boolean): void
        setSendMailAttach(sendMailAttach: boolean): void
        setShortMenuNames(shortMenuNames: boolean): void
        isShowControlCharacters(): boolean
        setShowControlCharacters(showControlCharacters: boolean): void
        setShowDiacritics(showDiacritics: boolean): void
        isShowFormatError(): boolean
        setShowFormatError(showFormatError: boolean): void
        isShowMarkupOpenSave(): boolean
        setShowMarkupOpenSave(showMarkupOpenSave: boolean): void
        setSmartCursoring(smartCursoring: boolean): void
        isSmartParaSelection(): boolean
        setSmartParaSelection(smartParaSelection: boolean): void
        isStoreRSIDOnSave(): boolean
        setStoreRSIDOnSave(storeRSIDOnSave: boolean): void
        setStrictFinalYaa(strictFinalYaa: boolean): void
        isStrictInitialAlefHamza(): boolean
        isUpdateFieldsAtPrint(): boolean
        isPrintXMLTag(): boolean
        setPrintXMLTag(printXMLTag: boolean): void
        isRTFInClipboard(): boolean
        getSaveInterval(): int
        setSaveInterval(saveInterval: int): void
        isSendMailAttach(): boolean
        isSequenceCheck(): boolean
        setSequenceCheck(sequenceCheck: boolean): void
        isShortMenuNames(): boolean
        isShowDiacritics(): boolean
        isSmartCursoring(): boolean
        isSmartCutPaste(): boolean
        setSmartCutPaste(smartCutPaste: boolean): void
        isSnapToGrid(): boolean
        setSnapToGrid(snapToGrid: boolean): void
        isSnapToShapes(): boolean
        setSnapToShapes(snapToShapes: boolean): void
        isStrictFinalYaa(): boolean
        isTabIndentKey(): boolean
        setTabIndentKey(tabIndentKey: boolean): void
        isTypeNReplace(): boolean
        setTypeNReplace(typeNReplace: boolean): void
        isWPDocNavKeys(): boolean
        setWPDocNavKeys(docNavKeys: boolean): void
        setWPHelp(help: boolean): void
        setWPHelpOptions(commandKeyHelp: boolean, docNavigationKeys: boolean, mouseSimulation: boolean, demoGuidance: boolean, demoSpeed: int, helpType: int): void
        setInputMode(include: boolean): void
        setPenMarksID(ID: string): void
        isShowReadabilityStatistics(): boolean
        setShowReadabilityStatistics(showReadabilityStatistics: boolean): void
        setStrictInitialAlefHamza(strictInitialAlefHamza: boolean): void
        isSuggestFromMainDictionaryOnly(): boolean
        setSuggestFromMainDictionaryOnly(suggestFromMainDictionaryOnly: boolean): void
        isSuggestSpellingCorrections(): boolean
        setSuggestSpellingCorrections(suggestSpellingCorrections: boolean): void
        isUseGermanSpellingReform(): boolean
        setUseGermanSpellingReform(useGermanSpellingReform: boolean): void
        setDocFieldShadingVisible(visible: boolean): void
        setUpdateFieldsWithTrackedChangesAtPrint(update: boolean): void
        isUpdateFieldsWithTrackedChangesAtPrint(): boolean
        isWarnBeforeSavingPrintingSendingMarkup(): boolean
        setWarnBeforeSavingPrintingSendingMarkup(warnBeforeSavingPrintingSendingMarkup: boolean): void
        setUpdateFieldsAtPrint(updateFieldsAtPrint: boolean): void
        isUpdateLinksAtOpen(): boolean
        setUpdateLinksAtOpen(updateLinksAtOpen: boolean): void
        isUpdateLinksAtPrint(): boolean
        setUpdateLinksAtPrint(updateLinksAtPrint: boolean): void
        isUseCharacterUnit(): boolean
        setUseCharacterUnit(useCharacterUnit: boolean): void
        isUseDiffDiacColor(): boolean
        setUseDiffDiacColor(useDiffDiacColor: boolean): void
        getVisualSelection(): int
        setVisualSelection(visualSelection: int): void
        getInputModeForDocField(): boolean
        setShowDocFieldMarks(visible: boolean): void
        isShowDocFieldMarks(): boolean
        isDocFieldShadingVisible(): boolean
        disableAccelerators(): void
        setCheckPenMarksUserName(value: boolean): void
    }
    export interface Options {
        setColor(ywColorIndex: int, what: int, setting: number[]): number[]
        getDefaultFilePath(path: int): string
        setShowErrorDialog(showFlag: boolean): boolean
        isShowErrorDialog(): boolean
        isAddBiDirectionalMarksWhenSavingTextFile(): boolean
        setAddBiDirectionalMarksWhenSavingTextFile(addBiDirectionalMarksWhenSavingTextFile: boolean): void
        isAutoFormatAsYouTypeFormatListItemBeginning(): boolean
        setAutoFormatAsYouTypeFormatListItemBeginning(autoFormatAsYouTypeFormatListItemBeginning: boolean): void
        isAutoFormatAsYouTypeReplaceFarEastDashes(): boolean
        setAutoFormatAsYouTypeReplaceFarEastDashes(autoFormatAsYouTypeReplaceFarEastDashes: boolean): void
        setLastEditPaneEnabled(enable: boolean): void
        getDeletedTextColor(): int
        setDeletedTextColor(deletedTextColor: int): void
        getDeletedTextMark(): int
        setDeletedTextMark(deletedTextMark: int): void
        getDiacriticColorVal(): int
        setDiacriticColorVal(diacriticColorVal: int): void
        isDisplayGridLines(): boolean
        setDisplayGridLines(displayGridLines: boolean): void
        isDisplayPasteOptions(): boolean
        setDisplayPasteOptions(displayPasteOptions: boolean): void
        isDisplaySmartTagButtons(): boolean
        getDocumentViewDirection(): int
        setDocumentViewDirection(documentViewDirection: int): void
        setEormatScanning(eormatScanning: boolean): void
        getGridDistanceVertical(): double
        setGridDistanceVertical(gridDistanceVertical: double): void
        getGridOriginHorizontal(): double
        setGridOriginHorizontal(gridOriginHorizontal: double): void
        getGridOriginVertical(): double
        setGridOriginVertical(gridOriginVertical: double): void
        isIgnoreMixedDigits(): boolean
        setIgnoreMixedDigits(ignoreMixedDigits: boolean): void
        isIgnoreUppercase(): boolean
        setIgnoreUppercase(ignoreUppercase: boolean): void
        isIMEAutomaticControl(): boolean
        setIMEAutomaticControl(automaticControl: boolean): void
        isInlineConversion(): boolean
        setInlineConversion(inlineConversion: boolean): void
        getInsertedTextColor(): int
        setInsertedTextColor(insertedTextColor: int): void
        getInsertedTextMark(): int
        setInsertedTextMark(insertedTextMark: int): void
        setINSKeyForPaste(keyForPaste: boolean): void
        getInterpretHighAnsi(): int
        setInterpretHighAnsi(interpretHighAnsi: int): void
        setLabelSmartTags(labelSmartTags: boolean): void
        isLocalNetworkFile(): boolean
        setLocalNetworkFile(localNetworkFile: boolean): void
        setMatchFuzzyByte(matchFuzzyByte: boolean): void
        isDisableFeaturesbyDefault(): boolean
        setDisableFeaturesbyDefault(disableFeaturesbyDefault: boolean): void
        setDisplaySmartTagButtons(displaySmartTagButtons: boolean): void
        isEnableMisusedWordsDictionary(): boolean
        setEnableMisusedWordsDictionary(enableMisusedWordsDictionary: boolean): void
        isEnvelopeFeederInstalled(): boolean
        setEnvelopeFeederInstalled(envelopeFeederInstalled: boolean): void
        getGridDistanceHorizontal(): double
        setGridDistanceHorizontal(gridDistanceHorizontal: double): void
        isHangulHanjaFastConversion(): boolean
        setHangulHanjaFastConversion(hangulHanjaFastConversion: boolean): void
        isIgnoreInternetAndFileAddresses(): boolean
        isMatchFuzzyIterationMark(): boolean
        setMatchFuzzyIterationMark(matchFuzzyIterationMark: boolean): void
        isEnableHangulHanjaRecentOrdering(): boolean
        setEnableHangulHanjaRecentOrdering(enableHangulHanjaRecentOrdering: boolean): void
        setIgnoreInternetAndFileAddresses(ignoreInternetAndFileAddresses: boolean): void
        setMatchFuzzyCase(matchFuzzyCase: boolean): void
        setMatchFuzzyDash(matchFuzzyDash: boolean): void
        isMatchFuzzyHiragana(): boolean
        setMatchFuzzyHiragana(matchFuzzyHiragana: boolean): void
        isMatchFuzzyKanji(): boolean
        setMatchFuzzyKanji(matchFuzzyKanji: boolean): void
        setMatchFuzzyKiKu(matchFuzzyKiKu: boolean): void
        isMatchFuzzyOldKana(): boolean
        setMatchFuzzyOldKana(matchFuzzyOldKana: boolean): void
        isMatchFuzzyPunctuation(): boolean
        setMatchFuzzyPunctuation(matchFuzzyPunctuation: boolean): void
        isMatchFuzzySmallKana(): boolean
        setMatchFuzzySmallKana(matchFuzzySmallKana: boolean): void
        isMatchFuzzySpace(): boolean
        setMatchFuzzySpace(matchFuzzySpace: boolean): void
        getMeasurementUnit(): int
        setMeasurementUnit(measurementUnit: int): void
        isPasteAdjustWordSpacing(): boolean
        isPasteMergeFromPPT(): boolean
        setPasteMergeFromPPT(pasteMergeFromPPT: boolean): void
        isPasteMergeFromXL(): boolean
        setPasteMergeFromXL(pasteMergeFromXL: boolean): void
        isPasteMergeLists(): boolean
        setPasteMergeLists(pasteMergeLists: boolean): void
        isPasteSmartCutPaste(): boolean
        setPasteSmartCutPaste(pasteSmartCutPaste: boolean): void
        getPictureWrapType(): int
        setPictureWrapType(pictureWrapType: int): void
        isPrintBackground(): boolean
        setPrintBackground(printBackground: boolean): void
        isPrintBackgrounds(): boolean
        setPrintBackgrounds(printBackgrounds: boolean): void
        setPrintHiddenText(printHiddenText: boolean): void
        isPrintDrawingObjects(): boolean
        setPrintDrawingObjects(printDrawingObjects: boolean): void
        isPrintFieldCodes(): boolean
        setPrintFieldCodes(printFieldCodes: boolean): void
        isPrintHiddenText(): boolean
        isMatchFuzzyProintedSoundMark(): boolean
        setMatchFuzzyProintedSoundMark(matchFuzzyProintedSoundMark: boolean): void
        getMultipleWordConversionsMode(): int
        setMultipleWordConversionsMode(multipleWordConversionsMode: int): void
        isOptimizeForWord97byDefault(): boolean
        setOptimizeForWord97byDefault(optimizeForWord97byDefault: boolean): void
        isPasteAdjustParagraphSpacing(): boolean
        setPasteAdjustParagraphSpacing(pasteAdjustParagraphSpacing: boolean): void
        isPasteAdjustTableFormatting(): boolean
        setPasteAdjustTableFormatting(pasteAdjustTableFormatting: boolean): void
        setPasteAdjustWordSpacing(pasteAdjustWordSpacing: boolean): void
        isPasteSmartStyleBehavior(): boolean
        setPasteSmartStyleBehavior(pasteSmartStyleBehavior: boolean): void
        isPrintEvenPagesInAscendingOrder(): boolean
        isPrintOddPagesInAscendingOrder(): boolean
        setPrintOddPagesInAscendingOrder(printOddPagesInAscendingOrder: boolean): void
        getRevisedPropertiesColor(): int
        setRevisedPropertiesColor(revisedPropertiesColor: int): void
        setPrintEvenPagesInAscendingOrder(printEvenPagesInAscendingOrder: boolean): void
        getRevisionsBalloonPrintOrientation(): int
        setRevisionsBalloonPrintOrientation(revisionsBalloonPrintOrientation: int): void
        getColor(what: int): int
        isWPHelp(): boolean
        setAllowAccentedUppercase(allowAccentedUppercase: boolean): void
        isAllowFastSave(): boolean
        setAllowFastSave(allowFastSave: boolean): void
        getArabicMode(): int
        setArabicMode(arabicMode: int): void
        getArabicNumeral(): int
        setArabicNumeral(arabicNumeral: int): void
        isBackgroundOpen(): boolean
        isBackgroundSave(): boolean
        isBlueScreen(): boolean
        setBlueScreen(blueScreen: boolean): void
        getCommentsColor(): int
        setCommentsColor(commentsColor: int): void
        isCreateBackup(): boolean
        setCreateBackup(createBackup: boolean): void
        getDefaultTray(): string
        setDefaultTray(defaultTray: string): void
        getDefaultTrayID(): int
        setDefaultTrayID(defaultTrayID: int): void
        isEnableSound(): boolean
        setEnableSound(enableSound: boolean): void
        isEormatScanning(): boolean
        getHebrewMode(): int
        setHebrewMode(hebrewMode: int): void
        isINSKeyForPaste(): boolean
        isLabelSmartTags(): boolean
        isMapPaperSize(): boolean
        setMapPaperSize(mapPaperSize: boolean): void
        isMatchFuzzyAY(): boolean
        setMatchFuzzyAY(matchFuzzyAY: boolean): void
        isMatchFuzzyBV(): boolean
        setMatchFuzzyBV(matchFuzzyBV: boolean): void
        isMatchFuzzyByte(): boolean
        isMatchFuzzyCase(): boolean
        isMatchFuzzyDash(): boolean
        isMatchFuzzyDZ(): boolean
        setMatchFuzzyDZ(matchFuzzyDZ: boolean): void
        isMatchFuzzyHF(): boolean
        setMatchFuzzyHF(matchFuzzyHF: boolean): void
        isMatchFuzzyKiKu(): boolean
        isMatchFuzzyTC(): boolean
        setMatchFuzzyTC(matchFuzzyTC: boolean): void
        isMatchFuzzyZJ(): boolean
        setMatchFuzzyZJ(matchFuzzyZJ: boolean): void
        getMonthNames(): int
        setMonthNames(monthNames: int): void
        isOvertype(): boolean
        setOvertype(overtype: boolean): void
        isPagination(): boolean
        setPagination(pagination: boolean): void
        getPictureEditor(): string
        setPictureEditor(pictureEditor: string): void
        isPrintComments(): boolean
        setPrintComments(printComments: boolean): void
        isPrintDraft(): boolean
        setPrintDraft(printDraft: boolean): void
        isPrintReverse(): boolean
        setPrintReverse(printReverse: boolean): void
        isAutoFormatAsYouTypeApplyBorders(): boolean
        setAutoFormatAsYouTypeApplyBorders(autoFormatAsYouTypeApplyBorders: boolean): void
        isAutoFormatAsYouTypeApplyBulletedLists(): boolean
        setAutoFormatAsYouTypeApplyBulletedLists(autoFormatAsYouTypeApplyBulletedLists: boolean): void
        isAutoFormatAsYouTypeApplyClosings(): boolean
        setAutoFormatAsYouTypeApplyClosings(autoFormatAsYouTypeApplyClosings: boolean): void
        isAutoFormatAsYouTypeApplyFirstIndents(): boolean
        setAutoFormatAsYouTypeApplyFirstIndents(autoFormatAsYouTypeApplyFirstIndents: boolean): void
        isAutoFormatAsYouTypeApplyHeadings(): boolean
        setAutoFormatAsYouTypeApplyHeadings(autoFormatAsYouTypeApplyHeadings: boolean): void
        isAutoFormatAsYouTypeApplyNumberedLists(): boolean
        setAutoFormatAsYouTypeApplyNumberedLists(autoFormatAsYouTypeApplyNumberedLists: boolean): void
        setAutoFormatAsYouTypeApplyTables(autoFormatAsYouTypeApplyTables: boolean): void
        isAutoFormatAsYouTypeAutoLetterWizard(): boolean
        setAutoFormatAsYouTypeAutoLetterWizard(autoFormatAsYouTypeAutoLetterWizard: boolean): void
        isAutoFormatAsYouTypeDefineStyles(): boolean
        setAutoFormatAsYouTypeDefineStyles(autoFormatAsYouTypeDefineStyles: boolean): void
        isAutoFormatAsYouTypeDeleteAutoSpaces(): boolean
        setAutoFormatAsYouTypeDeleteAutoSpaces(autoFormatAsYouTypeDeleteAutoSpaces: boolean): void
        isAutoFormatAsYouTypeInsertClosings(): boolean
        setAutoFormatAsYouTypeInsertClosings(autoFormatAsYouTypeInsertClosings: boolean): void
        setAutoFormatAsYouTypeInsertOvers(autoFormatAsYouTypeInsertOvers: boolean): void
        isAutoFormatAsYouTypeMatchParentheses(): boolean
        isAddControlCharacters(): boolean
        setAddControlCharacters(addControlCharacters: boolean): void
        isAddHebDoubleQuote(): boolean
        setAddHebDoubleQuote(addHebDoubleQuote: boolean): void
        isAllowAccentedUppercase(): boolean
        isAllowClickAndTypeMouse(): boolean
        isAllowDragAndDrop(): boolean
        setAllowDragAndDrop(allowDragAndDrop: boolean): void
        isAllowPixelUnits(): boolean
        setAllowPixelUnits(allowPixelUnits: boolean): void
        setAllowClickAndTypeMouse(allowClickAndTypeMouse: boolean): void
        isAllowCombinedAuxiliaryForms(): boolean
        setAllowCombinedAuxiliaryForms(allowCombinedAuxiliaryForms: boolean): void
        isAllowCompoundNounProcessing(): boolean
        setAllowCompoundNounProcessing(allowCompoundNounProcessing: boolean): void
        setAnimateScreenMovements(animateScreenMovements: boolean): void
        isApplyFarEastFontsToAscii(): boolean
        setApplyFarEastFontsToAscii(applyFarEastFontsToAscii: boolean): void
        isAutoFormatApplyBulletedLists(): boolean
        setAutoFormatApplyBulletedLists(autoFormatApplyBulletedLists: boolean): void
        isAutoFormatApplyFirstIndents(): boolean
        setAutoFormatApplyFirstIndents(autoFormatApplyFirstIndents: boolean): void
        isAutoFormatApplyHeadings(): boolean
        setAutoFormatApplyHeadings(autoFormatApplyHeadings: boolean): void
        isAutoFormatApplyOtherParas(): boolean
        setAutoFormatApplyOtherParas(autoFormatApplyOtherParas: boolean): void
        isAutoFormatAsYouTypeApplyDates(): boolean
        setAutoFormatAsYouTypeApplyDates(autoFormatAsYouTypeApplyDates: boolean): void
        isAutoFormatAsYouTypeApplyTables(): boolean
        isAutoFormatAsYouTypeInsertOvers(): boolean
        isAutoFormatDeleteAutoSpaces(): boolean
        setAutoFormatDeleteAutoSpaces(autoFormatDeleteAutoSpaces: boolean): void
        isAutoFormatMatchParentheses(): boolean
        setAutoFormatMatchParentheses(autoFormatMatchParentheses: boolean): void
        isAutoFormatPlainTextWordMail(): boolean
        setAutoFormatPlainTextWordMail(autoFormatPlainTextWordMail: boolean): void
        isAutoFormatPreserveStyles(): boolean
        setAutoFormatPreserveStyles(autoFormatPreserveStyles: boolean): void
        isAutoFormatReplaceFarEastDashes(): boolean
        isAutoFormatReplaceFractions(): boolean
        setAutoFormatReplaceFractions(autoFormatReplaceFractions: boolean): void
        isAutoFormatReplaceHyperlinks(): boolean
        setAutoFormatReplaceHyperlinks(autoFormatReplaceHyperlinks: boolean): void
        isAutoFormatReplaceOrdinals(): boolean
        setAutoFormatReplaceOrdinals(autoFormatReplaceOrdinals: boolean): void
        isAutoFormatReplaceQuotes(): boolean
        setAutoFormatReplaceQuotes(autoFormatReplaceQuotes: boolean): void
        isAllowReadingMode(): boolean
        setAllowReadingMode(allowReadingMode: boolean): void
        isAnimateScreenMovements(): boolean
        isAutoCreateNewDrawings(): boolean
        setAutoCreateNewDrawings(autoCreateNewDrawings: boolean): void
        isAutoFormatApplyLists(): boolean
        setAutoFormatApplyLists(autoFormatApplyLists: boolean): void
        isAutoKeyboardSwitching(): boolean
        setAutoKeyboardSwitching(autoKeyboardSwitching: boolean): void
        isAutoWordSelection(): boolean
        setAutoWordSelection(autoWordSelection: boolean): void
        setBackgroundOpen(backgroundOpen: boolean): void
        setBackgroundSave(backgroundSave: boolean): void
        getButtonFieldClicks(): int
        setButtonFieldClicks(buttonFieldClicks: int): void
        isCheckGrammarAsYouType(): boolean
        setCheckGrammarAsYouType(checkGrammarAsYouType: boolean): void
        isCheckHangulEndings(): boolean
        setCheckHangulEndings(checkHangulEndings: boolean): void
        isCheckSpellingAsYouType(): boolean
        applyTrackChangeSetting(setting: number[]): void
        isConfirmConversions(): boolean
        setConfirmConversions(confirmConversions: boolean): void
        getCursorMovement(): int
        setCursorMovement(cursorMovement: int): void
        getDefaultBorderColor(): int
        setDefaultBorderColor(defaultBorderColor: int): void
        getDefaultEPostageApp(): string
        setDefaultEPostageApp(defaultEPostageApp: string): void
        setDefaultFilePath(path: int, defaultFilePath: string): void
        getDefaultOpenFormat(): int
        setDefaultOpenFormat(defaultOpenFormat: int): void
        getDefaultTextEncoding(): int
        setDefaultTextEncoding(defaultTextEncoding: int): void
        setAutoFormatAsYouTypeMatchParentheses(autoFormatAsYouTypeMatchParentheses: boolean): void
        isAutoFormatAsYouTypeReplaceFractions(): boolean
        setAutoFormatAsYouTypeReplaceFractions(autoFormatAsYouTypeReplaceFractions: boolean): void
        isAutoFormatAsYouTypeReplaceHyperlinks(): boolean
        setAutoFormatAsYouTypeReplaceHyperlinks(autoFormatAsYouTypeReplaceHyperlinks: boolean): void
        isAutoFormatAsYouTypeReplaceOrdinals(): boolean
        setAutoFormatAsYouTypeReplaceOrdinals(autoFormatAsYouTypeReplaceOrdinals: boolean): void
        isAutoFormatAsYouTypeReplaceQuotes(): boolean
        setAutoFormatAsYouTypeReplaceQuotes(autoFormatAsYouTypeReplaceQuotes: boolean): void
        isAutoFormatAsYouTypeReplaceSymbols(): boolean
        setAutoFormatAsYouTypeReplaceSymbols(autoFormatAsYouTypeReplaceSymbols: boolean): void
        setAutoFormatReplaceFarEastDashes(autoFormatReplaceFarEastDashes: boolean): void
        isAutoFormatAsYouTypeReplacePlainTextEmphasis(): boolean
        setAutoFormatAsYouTypeReplacePlainTextEmphasis(autoFormatAsYouTypeReplacePlainTextEmphasis: boolean): void
        getDisableFeaturesIntroducedAfterbyDefault(): int
        setDisableFeaturesIntroducedAfterbyDefault(disableFeaturesIntroducedAfterbyDefault: int): void
        isAutoFormatReplacePlainTextEmphasis(): boolean
        setAutoFormatReplacePlainTextEmphasis(autoFormatReplacePlainTextEmphasis: boolean): void
        isAutoFormatReplaceSymbols(): boolean
        setAutoFormatReplaceSymbols(autoFormatReplaceSymbols: boolean): void
        isCheckGrammarWithSpelling(): boolean
        setCheckGrammarWithSpelling(checkGrammarWithSpelling: boolean): void
        setCheckSpellingAsYouType(checkSpellingAsYouType: boolean): void
        getapplyTrackChangeSetting(): number[]
        isConvertHighAnsiToFarEast(): boolean
        setConvertHighAnsiToFarEast(convertHighAnsiToFarEast: boolean): void
        isCtrlClickHyperlinkToOpen(): boolean
        setCtrlClickHyperlinkToOpen(ctrlClickHyperlinkToOpen: boolean): void
        getDefaultBorderColorIndex(): int
        setDefaultBorderColorIndex(defaultBorderColorIndex: int): void
        getDefaultBorderLineStyle(): int
        setDefaultBorderLineStyle(defaultBorderLineStyle: int): void
        getDefaultBorderLineWidth(): int
        setDefaultBorderLineWidth(defaultBorderLineWidth: int): void
        getDefaultHighlightColorIndex(): int
        setDefaultHighlightColorIndex(defaultHighlightColorIndex: int): void
        isPrintProperties(): boolean
        setPrintProperties(printProperties: boolean): void
        isPromptUpdateStyle(): boolean
        setPromptUpdateStyle(promptUpdateStyle: boolean): void
        isReplaceSelection(): boolean
        setReplaceSelection(replaceSelection: boolean): void
        getRevisedLinesColor(): int
        setRevisedLinesColor(revisedLinesColor: int): void
        getRevisedLinesMark(): int
        setRevisedLinesMark(revisedLinesMark: int): void
        getRevisedPropertiesMark(): int
        setRevisedPropertiesMark(revisedPropertiesMark: int): void
        setRTFInClipboard(inClipboard: boolean): void
        isSaveNormalPrompt(): boolean
        setSaveNormalPrompt(saveNormalPrompt: boolean): void
        isSavePropertiesPrompt(): boolean
        setSavePropertiesPrompt(savePropertiesPrompt: boolean): void
        setSendMailAttach(sendMailAttach: boolean): void
        setShortMenuNames(shortMenuNames: boolean): void
        isShowControlCharacters(): boolean
        setShowControlCharacters(showControlCharacters: boolean): void
        setShowDiacritics(showDiacritics: boolean): void
        isShowFormatError(): boolean
        setShowFormatError(showFormatError: boolean): void
        isShowMarkupOpenSave(): boolean
        setShowMarkupOpenSave(showMarkupOpenSave: boolean): void
        setSmartCursoring(smartCursoring: boolean): void
        isSmartParaSelection(): boolean
        setSmartParaSelection(smartParaSelection: boolean): void
        isStoreRSIDOnSave(): boolean
        setStoreRSIDOnSave(storeRSIDOnSave: boolean): void
        setStrictFinalYaa(strictFinalYaa: boolean): void
        isStrictInitialAlefHamza(): boolean
        isUpdateFieldsAtPrint(): boolean
        isPrintXMLTag(): boolean
        setPrintXMLTag(printXMLTag: boolean): void
        isRTFInClipboard(): boolean
        getSaveInterval(): int
        setSaveInterval(saveInterval: int): void
        isSendMailAttach(): boolean
        isSequenceCheck(): boolean
        setSequenceCheck(sequenceCheck: boolean): void
        isShortMenuNames(): boolean
        isShowDiacritics(): boolean
        isSmartCursoring(): boolean
        isSmartCutPaste(): boolean
        setSmartCutPaste(smartCutPaste: boolean): void
        isSnapToGrid(): boolean
        setSnapToGrid(snapToGrid: boolean): void
        isSnapToShapes(): boolean
        setSnapToShapes(snapToShapes: boolean): void
        isStrictFinalYaa(): boolean
        isTabIndentKey(): boolean
        setTabIndentKey(tabIndentKey: boolean): void
        isTypeNReplace(): boolean
        setTypeNReplace(typeNReplace: boolean): void
        isWPDocNavKeys(): boolean
        setWPDocNavKeys(docNavKeys: boolean): void
        setWPHelp(help: boolean): void
        setWPHelpOptions(commandKeyHelp: boolean, docNavigationKeys: boolean, mouseSimulation: boolean, demoGuidance: boolean, demoSpeed: int, helpType: int): void
        setInputMode(include: boolean): void
        setPenMarksID(ID: string): void
        isShowReadabilityStatistics(): boolean
        setShowReadabilityStatistics(showReadabilityStatistics: boolean): void
        setStrictInitialAlefHamza(strictInitialAlefHamza: boolean): void
        isSuggestFromMainDictionaryOnly(): boolean
        setSuggestFromMainDictionaryOnly(suggestFromMainDictionaryOnly: boolean): void
        isSuggestSpellingCorrections(): boolean
        setSuggestSpellingCorrections(suggestSpellingCorrections: boolean): void
        isUseGermanSpellingReform(): boolean
        setUseGermanSpellingReform(useGermanSpellingReform: boolean): void
        setDocFieldShadingVisible(visible: boolean): void
        setUpdateFieldsWithTrackedChangesAtPrint(update: boolean): void
        isUpdateFieldsWithTrackedChangesAtPrint(): boolean
        isWarnBeforeSavingPrintingSendingMarkup(): boolean
        setWarnBeforeSavingPrintingSendingMarkup(warnBeforeSavingPrintingSendingMarkup: boolean): void
        setUpdateFieldsAtPrint(updateFieldsAtPrint: boolean): void
        isUpdateLinksAtOpen(): boolean
        setUpdateLinksAtOpen(updateLinksAtOpen: boolean): void
        isUpdateLinksAtPrint(): boolean
        setUpdateLinksAtPrint(updateLinksAtPrint: boolean): void
        isUseCharacterUnit(): boolean
        setUseCharacterUnit(useCharacterUnit: boolean): void
        isUseDiffDiacColor(): boolean
        setUseDiffDiacColor(useDiffDiacColor: boolean): void
        getVisualSelection(): int
        setVisualSelection(visualSelection: int): void
        getInputModeForDocField(): boolean
        setShowDocFieldMarks(visible: boolean): void
        isShowDocFieldMarks(): boolean
        isDocFieldShadingVisible(): boolean
        disableAccelerators(): void
        setCheckPenMarksUserName(value: boolean): void
    }
    export interface OtherCorrectionsException {
        getName(): string
        delete(): void
        getIndex(): int
    }
    export interface OtherCorrectionsException {
        getName(): string
        delete(): void
        getIndex(): int
    }
    export interface OtherCorrectionsExceptions {
        add(name: string): word.OtherCorrectionsException
        getName(): string
        item(index: any): word.OtherCorrectionsException
        getCount(): int
        getArrayList(): any
    }
    export interface OtherCorrectionsExceptions {
        add(name: string): word.OtherCorrectionsException
        getName(): string
        item(index: any): word.OtherCorrectionsException
        getCount(): int
        getArrayList(): any
    }
    export interface Page {
        getWidth(): int
        getLeft(): int
        getTop(): int
        getHeight(): int
        getRectangles(): word.Rectangles
        getPage(): any
        getBreaks(): word.Breaks
    }
    export interface Page {
        getWidth(): int
        getLeft(): int
        getTop(): int
        getHeight(): int
        getRectangles(): word.Rectangles
        getPage(): any
        getBreaks(): word.Breaks
    }
    export interface PageBorder {
        setColor(colorType: int): void
        isVisible(): boolean
        setVisible(visible: boolean): void
        getColor(): int
        isInside(): boolean
        setArtStyle(style: int): void
        getArtStyle(): int
        setArtWidth(width: int): void
        getArtWidth(): int
        getPageBordersAttr(): any
        setColorIndex(index: int): void
        getColorIndex(): int
        setLineStyle(style: int): void
        getLineStyle(): int
        setLineWidth(width: int): void
        getLineWidth(): int
    }
    export interface PageBorder {
        setColor(colorType: int): void
        isVisible(): boolean
        setVisible(visible: boolean): void
        getColor(): int
        isInside(): boolean
        setArtStyle(style: int): void
        getArtStyle(): int
        setArtWidth(width: int): void
        getArtWidth(): int
        getPageBordersAttr(): any
        setColorIndex(index: int): void
        getColorIndex(): int
        setLineStyle(style: int): void
        getLineStyle(): int
        setLineWidth(width: int): void
        getLineWidth(): int
    }
    export interface PageNumber {
        delete(): void
        copy(): void
        getIndex(): int
        cut(): void
        select(): void
        getAlignment(): int
        setAlignment(alignment: int): void
    }
    export interface PageNumber {
        delete(): void
        copy(): void
        getIndex(): int
        cut(): void
        select(): void
        getAlignment(): int
        setAlignment(alignment: int): void
    }
    export interface PageNumbers {
        add(pageNumberAlignment: any, firstPage: any): word.PageNumber
        add(pageNumberAlignment: int, firstPage: boolean): word.PageNumber
        apply(): void
        item(index: int): word.PageNumber
        getCount(): int
        setIncludeChapterNumber(includeChapterNumber: boolean): void
        isIncludeChapterNumber(): boolean
        setNumberStyle(numberStyle: int): void
        getNumberStyle(): int
        setStartingNumber(startingNumber: int): void
        getStartingNumber(): int
        getChapterPageSeparator(): int
        setChapterPageSeparator(chapterPageSeparator: int): void
        isDoubleQuote(): boolean
        setDoubleQuote(doubleQuote: boolean): void
        getHeadingLevelForChapter(): int
        setHeadingLevelForChapter(headingLevelForChapter: int): void
        isRestartNumberingAtSection(): boolean
        setRestartNumberingAtSection(restartNumberingAtSection: boolean): void
        isShowFirstPageNumber(): boolean
        setShowFirstPageNumber(showFirstPageNumber: boolean): void
    }
    export interface PageNumbers {
        add(pageNumberAlignment: any, firstPage: any): word.PageNumber
        add(pageNumberAlignment: int, firstPage: boolean): word.PageNumber
        apply(): void
        item(index: int): word.PageNumber
        getCount(): int
        setIncludeChapterNumber(includeChapterNumber: boolean): void
        isIncludeChapterNumber(): boolean
        setNumberStyle(numberStyle: int): void
        getNumberStyle(): int
        setStartingNumber(startingNumber: int): void
        getStartingNumber(): int
        getChapterPageSeparator(): int
        setChapterPageSeparator(chapterPageSeparator: int): void
        isDoubleQuote(): boolean
        setDoubleQuote(doubleQuote: boolean): void
        getHeadingLevelForChapter(): int
        setHeadingLevelForChapter(headingLevelForChapter: int): void
        isRestartNumberingAtSection(): boolean
        setRestartNumberingAtSection(restartNumberingAtSection: boolean): void
        isShowFirstPageNumber(): boolean
        setShowFirstPageNumber(showFirstPageNumber: boolean): void
    }
    export interface Pages {
        item(index: int): word.Page
        getCount(): int
    }
    export interface Pages {
        item(index: int): word.Page
        getCount(): int
    }
    export interface PageSetup {
        apply(): void
        getOrientation(): int
        setVerticalAlignment(verticalAlignment: int): void
        getVerticalAlignment(): int
        setOrientation(orientation: int): void
        getDifferentFirstPageHeaderFooter(): int
        setDifferentFirstPageHeaderFooter(differentFirstPageHeaderFooter: int): void
        getOddAndEvenPagesHeaderFooter(): int
        setOddAndEvenPagesHeaderFooter(oddAndEvenPagesHeaderFooter: int): void
        setLineNumbering(lineNumbering: word.LineNumbering): void
        setAsTemplateDefault(): void
        setTopMargin(topMargin: double): void
        getTopMargin(): double
        getBookFoldPrintingSheets(): int
        setBookFoldPrintingSheets(bookFoldPrintingSheets: int): void
        isBookFoldPrinting(): boolean
        setBookFoldPrinting(bookFoldPrinting: boolean): void
        isBookFoldRevPrinting(): boolean
        setBookFoldRevPrinting(bookFoldRevPrinting: boolean): void
        getFooterDistance(): double
        setFooterDistance(footerDistance: double): void
        getHeaderDistance(): double
        setHeaderDistance(headerDistance: double): void
        getOtherPagesTray(): int
        setOtherPagesTray(otherPagesTray: int): void
        getSectionDirection(): int
        setSectionDirection(sectionDirection: int): void
        getSuppressEndnotes(): int
        setSuppressEndnotes(suppressEndnotes: int): void
        getBottomMargin(): double
        setBottomMargin(bottomMargin: double): void
        getCharsLine(): double
        setCharsLine(charsLine: double): void
        getFirstPageTray(): int
        setFirstPageTray(firstPageTray: int): void
        getGutter(): double
        setGutter(gutter: double): void
        getGutterPos(): int
        setGutterPos(gutterPos: int): void
        getGutterStyle(): int
        setGutterStyle(gutterStyle: int): void
        getLayoutMode(): int
        setLayoutMode(layoutMode: int): void
        getLeftMargin(): double
        setLeftMargin(leftMargin: double): void
        getLineNumbering(): word.LineNumbering
        getLinesPage(): double
        setLinesPage(linesPage: double): void
        getMirrorMargins(): int
        setMirrorMargins(mirrorMargins: int): void
        getPageHeight(): double
        setPageHeight(pageHeight: double): void
        getPageWidth(): double
        setPageWidth(pageWidth: double): void
        getPaperSize(): int
        setPaperSize(paperSize: int): void
        getRightMargin(): double
        setRightMargin(rightMargin: double): void
        getSectionStart(): int
        setSectionStart(sectionStart: int): void
        isShowGrid(): boolean
        setShowGrid(showGrid: boolean): void
        getTextColumns(): word.TextColumns
        setTextColumns(textColumns: word.TextColumns): void
        isTwoPagesOnOne(): boolean
        setTwoPagesOnOne(twoPagesOnOne: boolean): void
        togglePortrait(): void
    }
    export interface PageSetup {
        apply(): void
        getOrientation(): int
        setVerticalAlignment(verticalAlignment: int): void
        getVerticalAlignment(): int
        setOrientation(orientation: int): void
        getDifferentFirstPageHeaderFooter(): int
        setDifferentFirstPageHeaderFooter(differentFirstPageHeaderFooter: int): void
        getOddAndEvenPagesHeaderFooter(): int
        setOddAndEvenPagesHeaderFooter(oddAndEvenPagesHeaderFooter: int): void
        setLineNumbering(lineNumbering: word.LineNumbering): void
        setAsTemplateDefault(): void
        setTopMargin(topMargin: double): void
        getTopMargin(): double
        getBookFoldPrintingSheets(): int
        setBookFoldPrintingSheets(bookFoldPrintingSheets: int): void
        isBookFoldPrinting(): boolean
        setBookFoldPrinting(bookFoldPrinting: boolean): void
        isBookFoldRevPrinting(): boolean
        setBookFoldRevPrinting(bookFoldRevPrinting: boolean): void
        getFooterDistance(): double
        setFooterDistance(footerDistance: double): void
        getHeaderDistance(): double
        setHeaderDistance(headerDistance: double): void
        getOtherPagesTray(): int
        setOtherPagesTray(otherPagesTray: int): void
        getSectionDirection(): int
        setSectionDirection(sectionDirection: int): void
        getSuppressEndnotes(): int
        setSuppressEndnotes(suppressEndnotes: int): void
        getBottomMargin(): double
        setBottomMargin(bottomMargin: double): void
        getCharsLine(): double
        setCharsLine(charsLine: double): void
        getFirstPageTray(): int
        setFirstPageTray(firstPageTray: int): void
        getGutter(): double
        setGutter(gutter: double): void
        getGutterPos(): int
        setGutterPos(gutterPos: int): void
        getGutterStyle(): int
        setGutterStyle(gutterStyle: int): void
        getLayoutMode(): int
        setLayoutMode(layoutMode: int): void
        getLeftMargin(): double
        setLeftMargin(leftMargin: double): void
        getLineNumbering(): word.LineNumbering
        getLinesPage(): double
        setLinesPage(linesPage: double): void
        getMirrorMargins(): int
        setMirrorMargins(mirrorMargins: int): void
        getPageHeight(): double
        setPageHeight(pageHeight: double): void
        getPageWidth(): double
        setPageWidth(pageWidth: double): void
        getPaperSize(): int
        setPaperSize(paperSize: int): void
        getRightMargin(): double
        setRightMargin(rightMargin: double): void
        getSectionStart(): int
        setSectionStart(sectionStart: int): void
        isShowGrid(): boolean
        setShowGrid(showGrid: boolean): void
        getTextColumns(): word.TextColumns
        setTextColumns(textColumns: word.TextColumns): void
        isTwoPagesOnOne(): boolean
        setTwoPagesOnOne(twoPagesOnOne: boolean): void
        togglePortrait(): void
    }
    export interface Pane {
        close(): void
        getIndex(): int
        activate(): void
        getSelection(): word.Selection
        getDocument(): word.Document
        getPrevious(): word.Pane
        getNext(): word.Pane
        getPages(): word.Pages
        getView(): word.View
        getZooms(): word.Zooms
        getFrameset(): word.Frameset
        getVerticalPercentScrolled(): int
        setVerticalPercentScrolled(verticalPercentScrolled: int): void
        setHorizontalPercentScrolled(horizontalPercentScrolled: int): void
        isDisplayVerticalRuler(): boolean
        setDisplayVerticalRuler(displayVerticalRuler: boolean): void
        getMinimumFontSize(): int
        setMinimumFontSize(minimumFontSize: int): void
        isDisplayRulers(): boolean
        setDisplayRulers(displayRulers: boolean): void
        getBrowseWidth(): int
        autoScroll(velocity: int): void
        largeScroll(down: int, up: int, toRight: int, toLeft: int): void
        newFrameset(): void
        pageScroll(down: int, up: int): void
        smallScroll(down: int, up: int, toRight: int, toLeft: int): void
        tocInFrameset(): void
    }
    export interface Pane {
        close(): void
        getIndex(): int
        activate(): void
        getSelection(): word.Selection
        getDocument(): word.Document
        getPrevious(): word.Pane
        getNext(): word.Pane
        getPages(): word.Pages
        getView(): word.View
        getZooms(): word.Zooms
        getFrameset(): word.Frameset
        getVerticalPercentScrolled(): int
        setVerticalPercentScrolled(verticalPercentScrolled: int): void
        setHorizontalPercentScrolled(horizontalPercentScrolled: int): void
        isDisplayVerticalRuler(): boolean
        setDisplayVerticalRuler(displayVerticalRuler: boolean): void
        getMinimumFontSize(): int
        setMinimumFontSize(minimumFontSize: int): void
        isDisplayRulers(): boolean
        setDisplayRulers(displayRulers: boolean): void
        getBrowseWidth(): int
        autoScroll(velocity: int): void
        largeScroll(down: int, up: int, toRight: int, toLeft: int): void
        newFrameset(): void
        pageScroll(down: int, up: int): void
        smallScroll(down: int, up: int, toRight: int, toLeft: int): void
        tocInFrameset(): void
    }
    export interface Panes {
        add(splitVertical: double): word.Pane
        item(index: int): word.Pane
        getCount(): int
    }
    export interface Panes {
        add(splitVertical: double): word.Pane
        item(index: int): word.Pane
        getCount(): int
    }
    export interface Paragraph {
        next(countObj: any): word.Paragraph
        reset(): void
        previous(countObj: any): word.Paragraph
        getID(): string
        getRange(): word.Range
        setID(id: string): void
        setBorders(borders: word.Borders): void
        getFormat(): word.ParagraphFormat
        getBorders(): word.Borders
        getShading(): word.Shading
        setWordWrap(wordWrap: int): void
        setFormat(format: word.ParagraphFormat): void
        getAlignment(): int
        setAlignment(alignment: int): void
        getLeftIndent(): double
        setLeftIndent(leftIndent: double): void
        getRightIndent(): double
        setRightIndent(rightIndent: double): void
        getSpaceAfter(): double
        setSpaceAfter(spaceAfter: double): void
        getSpaceBefore(): double
        setSpaceBefore(spaceBefore: double): void
        closeUp(): void
        indent(): void
        openUp(): void
        outdent(): void
        space1(): void
        space15(): void
        space2(): void
        getStyle(): any
        setStyle(style: any): void
        getMParagraph(): any
        getCharacterUnitFirstLineIndent(): double
        setCharacterUnitFirstLineIndent(characterUnitFirstLineIndent: double): void
        getCharacterUnitLeftIndent(): double
        setCharacterUnitLeftIndent(characterUnitLeftIndent: double): void
        getCharacterUnitRightIndent(): double
        getAddSpaceBetweenFarEastAndAlpha(): int
        setAddSpaceBetweenFarEastAndAlpha(addSpaceBetweenFarEastAndAlpha: int): void
        getAddSpaceBetweenFarEastAndDigit(): int
        setAddSpaceBetweenFarEastAndDigit(addSpaceBetweenFarEastAndDigit: int): void
        getHalfWidthPunctuationOnTopOfLine(): int
        setHalfWidthPunctuationOnTopOfLine(halfWidthPunctuationOnTopOfLine: int): void
        getAutoAdjustRightIndent(): int
        setAutoAdjustRightIndent(autoAdjustRightIndent: int): void
        getBaseLineAlignment(): int
        setBaseLineAlignment(baseLineAlignment: int): void
        getDisableLineHeightGrid(): int
        setDisableLineHeightGrid(disableLineHeightGrid: int): void
        getFirstLineIndent(): double
        setFirstLineIndent(firstLineIndent: double): void
        getHangingPunctuation(): int
        setHangingPunctuation(hangingPunctuation: int): void
        getLineSpacingRule(): int
        setLineSpacingRule(lineSpacingRule: int): void
        getLineUnitBefore(): double
        setLineUnitBefore(lineUnitBefore: double): void
        getPageBreakBefore(): int
        setPageBreakBefore(pageBreakBefore: int): void
        getSpaceAfterAuto(): int
        setSpaceAfterAuto(spaceAfterAuto: int): void
        getSpaceBeforeAuto(): int
        setSpaceBeforeAuto(spaceBeforeAuto: int): void
        indentFirstLineCharWidth(count: int): void
        outlineDemoteToBody(): void
        isStyleSeparator(): boolean
        getDropCap(): word.DropCap
        getHyphenation(): int
        setHyphenation(hyphenation: int): void
        getKeepTogether(): int
        setKeepTogether(keepTogether: int): void
        getKeepWithNext(): int
        setKeepWithNext(keepWithNext: int): void
        getLineSpacing(): double
        setLineSpacing(lineSpacing: double): void
        getLineUnitAfter(): double
        setLineUnitAfter(lineUnitAfter: double): void
        getNoLineNumber(): int
        setNoLineNumber(noLineNumber: int): void
        getOutlineLevel(): int
        setOutlineLevel(outlineLevel: int): void
        getReadingOrder(): int
        setReadingOrder(readingOrder: int): void
        getTabStops(): word.TabStops
        setTabStops(tabStops: word.TabStops): void
        getWidowControl(): int
        setWidowControl(widowControl: int): void
        getWordWrap(): int
        indentCharWidth(count: int): void
        openOrCloseUp(): void
        outlineDemote(): void
        outlinePromote(): void
        resetAdvanceTo(): void
        selectNumber(): void
        setCharacterUnitRightIndent(characterUnitRightIndent: double): void
        getFarEastLineBreakControl(): int
        setFarEastLineBreakControl(farEastLineBreakControl: int): void
        tabHangingIndent(count: int): void
        tabIndent(count: int): void
    }
    export interface Paragraph {
        next(countObj: any): word.Paragraph
        reset(): void
        previous(countObj: any): word.Paragraph
        getID(): string
        getRange(): word.Range
        setID(id: string): void
        setBorders(borders: word.Borders): void
        getFormat(): word.ParagraphFormat
        getBorders(): word.Borders
        getShading(): word.Shading
        setWordWrap(wordWrap: int): void
        setFormat(format: word.ParagraphFormat): void
        getAlignment(): int
        setAlignment(alignment: int): void
        getLeftIndent(): double
        setLeftIndent(leftIndent: double): void
        getRightIndent(): double
        setRightIndent(rightIndent: double): void
        getSpaceAfter(): double
        setSpaceAfter(spaceAfter: double): void
        getSpaceBefore(): double
        setSpaceBefore(spaceBefore: double): void
        closeUp(): void
        indent(): void
        openUp(): void
        outdent(): void
        space1(): void
        space15(): void
        space2(): void
        getStyle(): any
        setStyle(style: any): void
        getMParagraph(): any
        getCharacterUnitFirstLineIndent(): double
        setCharacterUnitFirstLineIndent(characterUnitFirstLineIndent: double): void
        getCharacterUnitLeftIndent(): double
        setCharacterUnitLeftIndent(characterUnitLeftIndent: double): void
        getCharacterUnitRightIndent(): double
        getAddSpaceBetweenFarEastAndAlpha(): int
        setAddSpaceBetweenFarEastAndAlpha(addSpaceBetweenFarEastAndAlpha: int): void
        getAddSpaceBetweenFarEastAndDigit(): int
        setAddSpaceBetweenFarEastAndDigit(addSpaceBetweenFarEastAndDigit: int): void
        getHalfWidthPunctuationOnTopOfLine(): int
        setHalfWidthPunctuationOnTopOfLine(halfWidthPunctuationOnTopOfLine: int): void
        getAutoAdjustRightIndent(): int
        setAutoAdjustRightIndent(autoAdjustRightIndent: int): void
        getBaseLineAlignment(): int
        setBaseLineAlignment(baseLineAlignment: int): void
        getDisableLineHeightGrid(): int
        setDisableLineHeightGrid(disableLineHeightGrid: int): void
        getFirstLineIndent(): double
        setFirstLineIndent(firstLineIndent: double): void
        getHangingPunctuation(): int
        setHangingPunctuation(hangingPunctuation: int): void
        getLineSpacingRule(): int
        setLineSpacingRule(lineSpacingRule: int): void
        getLineUnitBefore(): double
        setLineUnitBefore(lineUnitBefore: double): void
        getPageBreakBefore(): int
        setPageBreakBefore(pageBreakBefore: int): void
        getSpaceAfterAuto(): int
        setSpaceAfterAuto(spaceAfterAuto: int): void
        getSpaceBeforeAuto(): int
        setSpaceBeforeAuto(spaceBeforeAuto: int): void
        indentFirstLineCharWidth(count: int): void
        outlineDemoteToBody(): void
        isStyleSeparator(): boolean
        getDropCap(): word.DropCap
        getHyphenation(): int
        setHyphenation(hyphenation: int): void
        getKeepTogether(): int
        setKeepTogether(keepTogether: int): void
        getKeepWithNext(): int
        setKeepWithNext(keepWithNext: int): void
        getLineSpacing(): double
        setLineSpacing(lineSpacing: double): void
        getLineUnitAfter(): double
        setLineUnitAfter(lineUnitAfter: double): void
        getNoLineNumber(): int
        setNoLineNumber(noLineNumber: int): void
        getOutlineLevel(): int
        setOutlineLevel(outlineLevel: int): void
        getReadingOrder(): int
        setReadingOrder(readingOrder: int): void
        getTabStops(): word.TabStops
        setTabStops(tabStops: word.TabStops): void
        getWidowControl(): int
        setWidowControl(widowControl: int): void
        getWordWrap(): int
        indentCharWidth(count: int): void
        openOrCloseUp(): void
        outlineDemote(): void
        outlinePromote(): void
        resetAdvanceTo(): void
        selectNumber(): void
        setCharacterUnitRightIndent(characterUnitRightIndent: double): void
        getFarEastLineBreakControl(): int
        setFarEastLineBreakControl(farEastLineBreakControl: int): void
        tabHangingIndent(count: int): void
        tabIndent(count: int): void
    }
    export interface ParagraphFormat {
        apply(): void
        reset(): void
        setBorders(borders: word.Borders): void
        getBorders(): word.Borders
        getShading(): word.Shading
        setWordWrap(wordWrap: int): void
        setFormat(newFormat: word.ParagraphFormat): void
        getMParagraphFormat(): any
        getAlignment(): int
        setAlignment(alignment: int): void
        getLeftIndent(): double
        setLeftIndent(leftIndent: double): void
        getRightIndent(): double
        setRightIndent(rightIndent: double): void
        getSpaceAfter(): double
        setSpaceAfter(spaceAfter: double): void
        getSpaceBefore(): double
        setSpaceBefore(spaceBefore: double): void
        getDuplicate(): word.ParagraphFormat
        closeUp(): void
        openUp(): void
        space1(): void
        space15(): void
        space2(): void
        getStyle(): any
        setStyle(style: any): void
        getCharacterUnitFirstLineIndent(): float
        setCharacterUnitFirstLineIndent(characterUnitFirstLineIndent: double): void
        getCharacterUnitLeftIndent(): double
        setCharacterUnitLeftIndent(characterUnitLeftIndent: double): void
        getCharacterUnitRightIndent(): double
        getAddSpaceBetweenFarEastAndAlpha(): int
        setAddSpaceBetweenFarEastAndAlpha(addSpaceBetweenFarEastAndAlpha: int): void
        getAddSpaceBetweenFarEastAndDigit(): int
        setAddSpaceBetweenFarEastAndDigit(addSpaceBetweenFarEastAndDigit: int): void
        getHalfWidthPunctuationOnTopOfLine(): int
        setHalfWidthPunctuationOnTopOfLine(halfWidthPunctuationOnTopOfLine: int): void
        getAutoAdjustRightIndent(): int
        setAutoAdjustRightIndent(autoAdjustRightIndent: int): void
        getBaseLineAlignment(): int
        setBaseLineAlignment(baseLineAlignment: int): void
        getDisableLineHeightGrid(): int
        setDisableLineHeightGrid(disableLineHeightGrid: int): void
        getFirstLineIndent(): double
        setFirstLineIndent(firstLineIndent: double): void
        getHangingPunctuation(): int
        setHangingPunctuation(hangingPunctuation: int): void
        getLineSpacingRule(): int
        setLineSpacingRule(lineSpacingRule: int): void
        getLineUnitBefore(): double
        setLineUnitBefore(lineUnitBefore: double): void
        getPageBreakBefore(): int
        setPageBreakBefore(pageBreakBefore: int): void
        getSpaceAfterAuto(): int
        setSpaceAfterAuto(spaceAfterAuto: int): void
        getSpaceBeforeAuto(): int
        setSpaceBeforeAuto(spaceBeforeAuto: int): void
        indentFirstLineCharWidth(count: int): void
        getTextboxTightWrap(): int
        setTextboxTightWrap(textboxTightWrap: int): void
        getHyphenation(): int
        setHyphenation(hyphenation: int): void
        getKeepTogether(): int
        setKeepTogether(keepTogether: int): void
        getKeepWithNext(): int
        setKeepWithNext(keepWithNext: int): void
        getLineSpacing(): double
        setLineSpacing(lineSpacing: double): void
        getLineUnitAfter(): double
        setLineUnitAfter(lineUnitAfter: double): void
        getNoLineNumber(): int
        setNoLineNumber(noLineNumber: int): void
        getOutlineLevel(): int
        setOutlineLevel(outlineLevel: int): void
        getReadingOrder(): int
        setReadingOrder(readingOrder: int): void
        getTabStops(): word.TabStops
        setTabStops(tabStops: word.TabStops): void
        getWidowControl(): int
        setWidowControl(widowControl: int): void
        getWordWrap(): int
        indentCharWidth(count: int): void
        openOrCloseUp(): void
        setCharacterUnitRightIndent(characterUnitRightIndent: double): void
        getFarEastLineBreakControl(): int
        setFarEastLineBreakControl(farEastLineBreakControl: int): void
        tabHangingIndent(count: int): void
        tabIndent(count: int): void
        setNoApply(flag: boolean): void
        checkMaxPoint(f1: float, f2: float, unit: int, negative: boolean): void
        checkMaxLine(f: float, negative: boolean): void
    }
    export interface ParagraphFormat {
        apply(): void
        reset(): void
        setBorders(borders: word.Borders): void
        getBorders(): word.Borders
        getShading(): word.Shading
        setWordWrap(wordWrap: int): void
        setFormat(newFormat: word.ParagraphFormat): void
        getMParagraphFormat(): any
        getAlignment(): int
        setAlignment(alignment: int): void
        getLeftIndent(): double
        setLeftIndent(leftIndent: double): void
        getRightIndent(): double
        setRightIndent(rightIndent: double): void
        getSpaceAfter(): double
        setSpaceAfter(spaceAfter: double): void
        getSpaceBefore(): double
        setSpaceBefore(spaceBefore: double): void
        getDuplicate(): word.ParagraphFormat
        closeUp(): void
        openUp(): void
        space1(): void
        space15(): void
        space2(): void
        getStyle(): any
        setStyle(style: any): void
        getCharacterUnitFirstLineIndent(): float
        setCharacterUnitFirstLineIndent(characterUnitFirstLineIndent: double): void
        getCharacterUnitLeftIndent(): double
        setCharacterUnitLeftIndent(characterUnitLeftIndent: double): void
        getCharacterUnitRightIndent(): double
        getAddSpaceBetweenFarEastAndAlpha(): int
        setAddSpaceBetweenFarEastAndAlpha(addSpaceBetweenFarEastAndAlpha: int): void
        getAddSpaceBetweenFarEastAndDigit(): int
        setAddSpaceBetweenFarEastAndDigit(addSpaceBetweenFarEastAndDigit: int): void
        getHalfWidthPunctuationOnTopOfLine(): int
        setHalfWidthPunctuationOnTopOfLine(halfWidthPunctuationOnTopOfLine: int): void
        getAutoAdjustRightIndent(): int
        setAutoAdjustRightIndent(autoAdjustRightIndent: int): void
        getBaseLineAlignment(): int
        setBaseLineAlignment(baseLineAlignment: int): void
        getDisableLineHeightGrid(): int
        setDisableLineHeightGrid(disableLineHeightGrid: int): void
        getFirstLineIndent(): double
        setFirstLineIndent(firstLineIndent: double): void
        getHangingPunctuation(): int
        setHangingPunctuation(hangingPunctuation: int): void
        getLineSpacingRule(): int
        setLineSpacingRule(lineSpacingRule: int): void
        getLineUnitBefore(): double
        setLineUnitBefore(lineUnitBefore: double): void
        getPageBreakBefore(): int
        setPageBreakBefore(pageBreakBefore: int): void
        getSpaceAfterAuto(): int
        setSpaceAfterAuto(spaceAfterAuto: int): void
        getSpaceBeforeAuto(): int
        setSpaceBeforeAuto(spaceBeforeAuto: int): void
        indentFirstLineCharWidth(count: int): void
        getTextboxTightWrap(): int
        setTextboxTightWrap(textboxTightWrap: int): void
        getHyphenation(): int
        setHyphenation(hyphenation: int): void
        getKeepTogether(): int
        setKeepTogether(keepTogether: int): void
        getKeepWithNext(): int
        setKeepWithNext(keepWithNext: int): void
        getLineSpacing(): double
        setLineSpacing(lineSpacing: double): void
        getLineUnitAfter(): double
        setLineUnitAfter(lineUnitAfter: double): void
        getNoLineNumber(): int
        setNoLineNumber(noLineNumber: int): void
        getOutlineLevel(): int
        setOutlineLevel(outlineLevel: int): void
        getReadingOrder(): int
        setReadingOrder(readingOrder: int): void
        getTabStops(): word.TabStops
        setTabStops(tabStops: word.TabStops): void
        getWidowControl(): int
        setWidowControl(widowControl: int): void
        getWordWrap(): int
        indentCharWidth(count: int): void
        openOrCloseUp(): void
        setCharacterUnitRightIndent(characterUnitRightIndent: double): void
        getFarEastLineBreakControl(): int
        setFarEastLineBreakControl(farEastLineBreakControl: int): void
        tabHangingIndent(count: int): void
        tabIndent(count: int): void
        setNoApply(flag: boolean): void
        checkMaxPoint(f1: float, f2: float, unit: int, negative: boolean): void
        checkMaxLine(f: float, negative: boolean): void
    }
    export interface Paragraphs {
        add(range: any): word.Paragraph
        getFirst(): word.Paragraph
        delete(index: int): void
        reset(): void
        item(index: int): word.Paragraph
        getLast(): word.Paragraph
        getCount(): int
        setBorders(borders: word.Borders): void
        getFormat(): word.ParagraphFormat
        getBorders(): word.Borders
        getShading(): word.Shading
        setWordWrap(wordWrap: int): void
        setFormat(newFormat: word.ParagraphFormat): void
        getAlignment(): int
        setAlignment(alignment: int): void
        getLeftIndent(): double
        setLeftIndent(leftIndent: double): void
        getRightIndent(): double
        setRightIndent(rightIndent: double): void
        getSpaceAfter(): double
        setSpaceAfter(spaceAfter: double): void
        getSpaceBefore(): double
        setSpaceBefore(spaceBefore: double): void
        closeUp(): void
        indent(): void
        openUp(): void
        outdent(): void
        space1(): void
        space15(): void
        space2(): void
        getStyle(): any
        setStyle(style: any): void
        getCharacterUnitFirstLineIndent(): float
        setCharacterUnitFirstLineIndent(characterUnitFirstLineIndent: double): void
        getCharacterUnitLeftIndent(): double
        setCharacterUnitLeftIndent(characterUnitLeftIndent: double): void
        getCharacterUnitRightIndent(): double
        getAddSpaceBetweenFarEastAndAlpha(): int
        setAddSpaceBetweenFarEastAndAlpha(addSpaceBetweenFarEastAndAlpha: int): void
        getAddSpaceBetweenFarEastAndDigit(): int
        setAddSpaceBetweenFarEastAndDigit(addSpaceBetweenFarEastAndDigit: int): void
        getHalfWidthPunctuationOnTopOfLine(): int
        setHalfWidthPunctuationOnTopOfLine(halfWidthPunctuationOnTopOfLine: int): void
        getAutoAdjustRightIndent(): int
        setAutoAdjustRightIndent(autoAdjustRightIndent: int): void
        getBaseLineAlignment(): int
        setBaseLineAlignment(baseLineAlignment: int): void
        getDisableLineHeightGrid(): int
        setDisableLineHeightGrid(disableLineHeightGrid: int): void
        getFirstLineIndent(): double
        setFirstLineIndent(firstLineIndent: double): void
        getHangingPunctuation(): int
        setHangingPunctuation(hangingPunctuation: int): void
        getLineSpacingRule(): int
        setLineSpacingRule(lineSpacingRule: int): void
        getLineUnitBefore(): double
        setLineUnitBefore(lineUnitBefore: double): void
        getPageBreakBefore(): int
        setPageBreakBefore(pageBreakBefore: int): void
        getSpaceAfterAuto(): int
        setSpaceAfterAuto(spaceAfterAuto: int): void
        getSpaceBeforeAuto(): int
        setSpaceBeforeAuto(spaceBeforeAuto: int): void
        indentFirstLineCharWidth(count: int): void
        outlineDemoteToBody(): void
        getHyphenation(): int
        setHyphenation(hyphenation: int): void
        getKeepTogether(): int
        setKeepTogether(keepTogether: int): void
        getKeepWithNext(): int
        setKeepWithNext(keepWithNext: int): void
        getLineSpacing(): double
        setLineSpacing(lineSpacing: double): void
        getLineUnitAfter(): double
        setLineUnitAfter(lineUnitAfter: double): void
        getNoLineNumber(): int
        setNoLineNumber(noLineNumber: int): void
        getOutlineLevel(): int
        setOutlineLevel(outlineLevel: int): void
        getReadingOrder(): int
        setReadingOrder(readingOrder: int): void
        getTabStops(): word.TabStops
        setTabStops(tabStops: word.TabStops): void
        getWidowControl(): int
        setWidowControl(widowControl: int): void
        getWordWrap(): int
        indentCharWidth(count: int): void
        getMParagraphs(): any
        openOrCloseUp(): void
        outlineDemote(): void
        outlinePromote(): void
        setCharacterUnitRightIndent(characterUnitRightIndent: double): void
        getFarEastLineBreakControl(): int
        setFarEastLineBreakControl(farEastLineBreakControl: int): void
        tabHangingIndent(count: int): void
        tabIndent(count: int): void
        decreaseSpacing(): void
        increaseSpacing(): void
    }
    export interface Paragraphs {
        add(range: any): word.Paragraph
        getFirst(): word.Paragraph
        delete(index: int): void
        reset(): void
        item(index: int): word.Paragraph
        getLast(): word.Paragraph
        getCount(): int
        setBorders(borders: word.Borders): void
        getFormat(): word.ParagraphFormat
        getBorders(): word.Borders
        getShading(): word.Shading
        setWordWrap(wordWrap: int): void
        setFormat(newFormat: word.ParagraphFormat): void
        getAlignment(): int
        setAlignment(alignment: int): void
        getLeftIndent(): double
        setLeftIndent(leftIndent: double): void
        getRightIndent(): double
        setRightIndent(rightIndent: double): void
        getSpaceAfter(): double
        setSpaceAfter(spaceAfter: double): void
        getSpaceBefore(): double
        setSpaceBefore(spaceBefore: double): void
        closeUp(): void
        indent(): void
        openUp(): void
        outdent(): void
        space1(): void
        space15(): void
        space2(): void
        getStyle(): any
        setStyle(style: any): void
        getCharacterUnitFirstLineIndent(): float
        setCharacterUnitFirstLineIndent(characterUnitFirstLineIndent: double): void
        getCharacterUnitLeftIndent(): double
        setCharacterUnitLeftIndent(characterUnitLeftIndent: double): void
        getCharacterUnitRightIndent(): double
        getAddSpaceBetweenFarEastAndAlpha(): int
        setAddSpaceBetweenFarEastAndAlpha(addSpaceBetweenFarEastAndAlpha: int): void
        getAddSpaceBetweenFarEastAndDigit(): int
        setAddSpaceBetweenFarEastAndDigit(addSpaceBetweenFarEastAndDigit: int): void
        getHalfWidthPunctuationOnTopOfLine(): int
        setHalfWidthPunctuationOnTopOfLine(halfWidthPunctuationOnTopOfLine: int): void
        getAutoAdjustRightIndent(): int
        setAutoAdjustRightIndent(autoAdjustRightIndent: int): void
        getBaseLineAlignment(): int
        setBaseLineAlignment(baseLineAlignment: int): void
        getDisableLineHeightGrid(): int
        setDisableLineHeightGrid(disableLineHeightGrid: int): void
        getFirstLineIndent(): double
        setFirstLineIndent(firstLineIndent: double): void
        getHangingPunctuation(): int
        setHangingPunctuation(hangingPunctuation: int): void
        getLineSpacingRule(): int
        setLineSpacingRule(lineSpacingRule: int): void
        getLineUnitBefore(): double
        setLineUnitBefore(lineUnitBefore: double): void
        getPageBreakBefore(): int
        setPageBreakBefore(pageBreakBefore: int): void
        getSpaceAfterAuto(): int
        setSpaceAfterAuto(spaceAfterAuto: int): void
        getSpaceBeforeAuto(): int
        setSpaceBeforeAuto(spaceBeforeAuto: int): void
        indentFirstLineCharWidth(count: int): void
        outlineDemoteToBody(): void
        getHyphenation(): int
        setHyphenation(hyphenation: int): void
        getKeepTogether(): int
        setKeepTogether(keepTogether: int): void
        getKeepWithNext(): int
        setKeepWithNext(keepWithNext: int): void
        getLineSpacing(): double
        setLineSpacing(lineSpacing: double): void
        getLineUnitAfter(): double
        setLineUnitAfter(lineUnitAfter: double): void
        getNoLineNumber(): int
        setNoLineNumber(noLineNumber: int): void
        getOutlineLevel(): int
        setOutlineLevel(outlineLevel: int): void
        getReadingOrder(): int
        setReadingOrder(readingOrder: int): void
        getTabStops(): word.TabStops
        setTabStops(tabStops: word.TabStops): void
        getWidowControl(): int
        setWidowControl(widowControl: int): void
        getWordWrap(): int
        indentCharWidth(count: int): void
        getMParagraphs(): any
        openOrCloseUp(): void
        outlineDemote(): void
        outlinePromote(): void
        setCharacterUnitRightIndent(characterUnitRightIndent: double): void
        getFarEastLineBreakControl(): int
        setFarEastLineBreakControl(farEastLineBreakControl: int): void
        tabHangingIndent(count: int): void
        tabIndent(count: int): void
        decreaseSpacing(): void
        increaseSpacing(): void
    }
    export interface PenMarkManager {
        isSelected(): boolean
        isVisible(): boolean
        setVisible(hidden: boolean): void
        select(select: boolean): void
        setLineWidth(width: int): void
        getLineWidth(): int
        stopMark(): void
        deletePenMarksByUserName(userName: string): void
        deletePenMarksByID(ID: string): void
        beginMark(): void
        setPenType(type: int): void
        getPenType(): int
        setRubberMode(mode: boolean): void
        isRubberMode(): boolean
        setLineColor(red: int, green: int, blue: int): void
        setLineColor(color: any): void
        getLineColor(): any
    }
    export interface PenMarkManager {
        isSelected(): boolean
        isVisible(): boolean
        setVisible(hidden: boolean): void
        select(select: boolean): void
        setLineWidth(width: int): void
        getLineWidth(): int
        stopMark(): void
        deletePenMarksByUserName(userName: string): void
        deletePenMarksByID(ID: string): void
        beginMark(): void
        setPenType(type: int): void
        getPenType(): int
        setRubberMode(mode: boolean): void
        isRubberMode(): boolean
        setLineColor(red: int, green: int, blue: int): void
        setLineColor(color: any): void
        getLineColor(): any
    }
    export interface PictureFormat {
        getBrightness(): double
        setBrightness(brightness: double): void
        getTransparencyColor(): int
        setTransparencyColor(transparencyColor: int): void
        getTransparentBackground(): int
        setTransparentBackground(transparentBackground: int): void
        incrementBrightness(increment: double): void
        incrementContrast(increment: double): void
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
        incrementBrightness(increment: double): void
        incrementContrast(increment: double): void
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
    export interface PrintDialog {
        execute(): void
        getNumCopies(): int
    }
    export interface PrintDialog {
        execute(): void
        getNumCopies(): int
    }
    export interface ProofreadingErrors {
        getType(): int
        item(index: int): word.Range
        getCount(): int
    }
    export interface ProofreadingErrors {
        getType(): int
        item(index: int): word.Range
        getCount(): int
    }
    export interface Range {
        next(unitObj: any, countObj: any): word.Range
        getFields(): word.Fields
        delete(unit: any, count: any): int
        copy(): void
        expand(unit: any): int
        sort(excludeHeader: any, fieldNumber: any, sortFieldType: any, sortOrder: any, fieldNumber2: any, sortFieldType2: any, sortOrder2: any, fieldNumber3: any, sortFieldType3: any, sortOrder3: any, sortColumn: any, separator: any, caseSensitive: any, bidiSort: any, ignoreThe: any, ignoreKashida: any, ignoreDiacritics: any, ignoreHe: any, languageID: any): void
        previous(unitObj: any, countObj: any): word.Range
        getTable(): word.Table
        getID(): string
        inRange(range: word.Range): boolean
        isEqual(range: word.Range): boolean
        getFont(): word.Font
        setKana(kana: int): void
        getKana(): int
        getOrientation(): int
        refresh(word: any): void
        checkGrammar(): void
        checkSpelling(customDictionary: any, ignoreUppercase: any, alwaysSuggest: any, customDictionary2: any, customDictionary3: any, customDictionary4: any, customDictionary5: any, customDictionary6: any, customDictionary7: any, customDictionary8: any, customDictionary9: any, customDictionary10: any): void
        getSpellingSuggestions(customDictionary: any, ignoreUppercase: any, mainDictionary: any, suggestionMode: any, customDictionary2: any, customDictionary3: any, customDictionary4: any, customDictionary5: any, customDictionary6: any, customDictionary7: any, customDictionary8: any, customDictionary9: any, customDictionary10: any): word.SpellingSuggestions
        lookupNameProperties(): void
        move(unitObj: any, countObj: any): int
        getFind(): word.Find
        getText(): string
        moveLeft(unitObj: int, countObj: int, extendObj: int): int
        moveLeft(unitObj: any, countObj: any, extendObj: any): int
        moveLeft(unit: int, count: any, extend: any): int
        paste(): void
        setFont(font: word.Font): void
        cut(): void
        getMRange(): any
        getDocument(): word.Document
        moveRight(unitObj: any, countObj: any, extendObj: any): int
        moveRight(unitObj: int, countObj: int, extendObj: int): int
        moveRight(unit: int, count: any, extend: any): int
        insertParagraph(): void
        getFrames(): word.Frames
        getTables(): word.Tables
        setText(text: any): void
        setText(text: string): void
        getStart(): int
        setStart(start: int): void
        setEnd(end: long): void
        setEnd(end: int): void
        getEnd(): int
        select(): void
        insertAfter(text: string): void
        setEmptyObject(isEmpty: boolean): void
        getStoryType(): int
        getPageSetup(): word.PageSetup
        isProtect(): void
        getBookmarks(): word.Bookmarks
        getParagraphs(): word.Paragraphs
        setID(id: string): void
        getCells(): word.Cells
        insertCaption(label: any, title: any, titleAutoText: any, position: any, excludeLabel: any): void
        getSections(): word.Sections
        setBorders(borders: word.Borders): void
        getBorders(): word.Borders
        checkIsNull(): void
        getShading(): word.Shading
        setParagraphFormat(paragraphFormat: word.ParagraphFormat): void
        getParagraphFormat(): word.ParagraphFormat
        computeStatistics(statistic: int): int
        setGrammarChecked(grammarChecked: boolean): void
        getGrammaticalErrors(): word.ProofreadingErrors
        isLanguageDetected(): boolean
        setLanguageDetected(languageDetected: boolean): void
        getReadabilityStatistics(): word.ReadabilityStatistics
        setBold(bold: int): void
        getBold(): int
        pasteAndFormat(type: int): void
        tcscConverter(ywTCSCConverterDirection: int, commonTerms: boolean, useVariants: boolean): void
        setTwoLinesInOne(twoLinesInOne: int): void
        setFitTextWidth(fitTextWidth: double): void
        getFitTextWidth(): double
        getShapeRange(): word.ShapeRange
        pasteSpecial(iconIndex: any, link: any, placement: any, displayAsIcon: any, dataType: any, iconFileName: any, iconLabel: any): void
        setOrientation(orientation: int): void
        convertToTable(separator: any, numRowsObj: any, numColumnsObj: any, initialColumnWidthObj: any, formatObj: any, applyBordersObj: any, applyShadingObj: any, applyFontObj: any, applyColorObj: any, applyHeadingRowsObj: any, applyLastRowObj: any, applyFirstColumnObj: any, applyLastColumnObj: any, autoFitObj: any, autoFitBehaviorObj: any, defaultTableBehaviorObj: any): word.Table
        getNoProofing(): int
        setNoProofing(noProofing: int): void
        getCharacterWidth(): int
        insertParagraphAfter(): void
        insertParagraphBefore(): void
        selectWithoutScroll(): void
        getLanguageIDFarEast(): int
        setLanguageIDFarEast(languageIDFarEast: int): void
        getLanguageIDOther(): int
        setLanguageIDOther(languageIDOther: int): void
        setBoldBi(boldBi: int): void
        getBoldBi(): int
        setEmphasisMark(emphasisMark: int): void
        getEmphasisMark(): int
        setItalic(italic: int): void
        getItalic(): int
        setItalicBi(italicBi: int): void
        getItalicBi(): int
        setUnderline(underline: int): void
        getUnderline(): int
        insertAutoText(): void
        insertSymbol(characterNumber: int, font: any, unicode: any, bias: any): void
        isDisableCharacterSpaceGrid(): boolean
        setDisableCharacterSpaceGrid(disableCharacterSpaceGrid: boolean): void
        getColumns(): word.Columns
        getDuplicate(): word.Range
        getRows(): word.Rows
        getXML(dataOnly: boolean): string
        getCase(): int
        setCase(case1: int): void
        getInlineShapes(): word.InlineShapes
        insertFile(fileName: string, range: any, confirmConversions: any, link: any, attachment: any): void
        getRevisions(): word.Revisions
        getScripts(): office.Scripts
        getSentences(): word.Sentences
        getSmartTags(): word.SmartTags
        getSubdocuments(): word.Subdocuments
        getXMLNodes(): word.XMLNodes
        getWords(): word.Words
        getStyle(): any
        setRange(start: int, end: int): void
        setStyle(style: any): void
        goToNext(what: int): word.Range
        collapse(direction: any): void
        isSpellingChecked(): boolean
        setSpellingChecked(spellingChecked: boolean): void
        getSpellingErrors(): word.ProofreadingErrors
        getEditors(): word.Editors
        goTo(what: any, which: any, count: any, name: any): word.Range
        getLanguageID(): int
        setLanguageID(languageID: int): void
        autoFormat(): void
        getCharacters(): word.Characters
        getComments(): word.Comments
        detectLanguage(): void
        getEndnotes(): word.Endnotes
        getFootnotes(): word.Footnotes
        getFormFields(): word.FormFields
        isGrammarChecked(): boolean
        getHTMLDivisions(): word.HTMLDivisions
        getHyperlinks(): word.Hyperlinks
        getListParagraphs(): word.ListParagraphs
        goToPrevious(what: int): word.Range
        getEndnoteOptions(): word.EndnoteOptions
        getEnhMetaFileBits(): any
        getFootnoteOptions(): word.FootnoteOptions
        getNextStoryRange(): word.Range
        getParagraphStyle(): any
        getPreviousBookmarkID(): int
        getTopLevelTables(): word.Tables
        getTabStops(): word.TabStops
        setTabStops(tabStops: word.TabStops): void
        modifyEnclosure(style: any, symbol: any, enclosedText: any): void
        checkUpOffset(word: any): boolean
        moveEndUntil(Cset: any, count: any): int
        moveEndWhile(Cset: any, count: any): int
        moveStart(unit: any, count: any): int
        moveStartUntil(Cset: any, count: any): int
        moveStartWhile(Cset: any, count: any): int
        moveUntil(Cset: any, count: any): int
        moveWhile(Cset: any, count: any): int
        nextSubdocument(): void
        pasteAppendTable(): void
        pasteExcelTable(linkedToExcel: boolean, wordFormatting: boolean, rtf: boolean): void
        phoneticGuide(text: string, alignment: int, raise: int, fontSize: int, fontName: string): void
        sortAscending(): void
        sortDescending(): void
        subscribeTo(edition: string, format: int): void
        wholeStory(): void
        moveStartInTable(count: int): void
        moveEndInTable(count: int): void
        getTextAsHtml(charsetName: string): string
        getTextAsHtml(): string
        getTextAsMht(): string
        isEndOfRowMark(): boolean
        isEmptyObject(): boolean
        getBookmarkID(): int
        getInformation(type: int): any
        getListFormat(): word.ListFormat
        getStoryLength(): int
        getSynonymInfo(): word.SynonymInfo
        getXMLParentNode(): word.XMLNode
        getFormattedText(): word.Range
        setFormattedText(formattedText: word.Range): void
        isShowAll(): boolean
        setShowAll(showAll: boolean): void
        getTwoLinesInOne(): int
        calculate(): double
        checkSynonyms(): void
        copyAsPicture(): void
        createPublisher(edition: any, containsPICT: boolean, containsRTF: boolean, containsText: boolean): void
        insertBefore(text: string): void
        insertBreak(type: any): void
        insertDatabase(format: any, style: any, linkToSource: any, connection: any, SQLStatement: any, SQLStatement1: any, passwordDocument: any, PasswordTemplate: any, writePasswordDocument: any, writePasswordTemplate: any, dataSource: any, from: any, to: any, includeFields: any): void
        insertDateTime(dateTimeFormat: any, insertAsField: any, insertAsFullWidth: any, dateLanguage: any, calendarType: any): void
        insertFile1(fileName: string, range: any, confirmConversions: any, link: any, attachment: any): boolean
        insertXML(text: string, transform: any): void
        stopScreenUpdating(): void
        startScreenUpdating(): void
        setCharacterWidth(characterWidth: int): void
        isCombineCharacters(): boolean
        setCombineCharacters(combineCharacters: boolean): void
        getHighlightColorIndex(): int
        setHighlightColorIndex(highlightColorIndex: int): void
        getHorizontalInVertical(): int
        setHorizontalInVertical(horizontalInVertical: int): void
        translateToMSText(text: string): string
        getTextRetrievalMode(): word.TextRetrievalMode
        setTextRetrievalMode(textRetrievalMode: word.TextRetrievalMode): void
        convertHangulAndHanja(conversionsMode: any, fastConversion: any, checkHangulEnding: any, enableRecentOrdering: any, customDictionary: any): void
        goToEditableRange(editorID: any): word.Range
        insertCrossReference(referenceType: any, referenceKind: int, referenceItem: any, insertAsHyperlink: any, includePosition: any, separateNumbers: any, separatorString: any): void
        pasteAsNestedTable(): void
        previousSubdocument(): void
        endOf(unit: any, extend: any): int
        doStart(): void
        doEnd(): void
        inStory(range: word.Range): boolean
        moveEnd(unit: int, count: int): int
        moveEnd(unitObj: any, countObj: any): int
        relocate(direction: int): void
        startOf(unit: any, extend: any): int
    }
    export interface Range {
        next(unitObj: any, countObj: any): word.Range
        getFields(): word.Fields
        delete(unit: any, count: any): int
        copy(): void
        expand(unit: any): int
        sort(excludeHeader: any, fieldNumber: any, sortFieldType: any, sortOrder: any, fieldNumber2: any, sortFieldType2: any, sortOrder2: any, fieldNumber3: any, sortFieldType3: any, sortOrder3: any, sortColumn: any, separator: any, caseSensitive: any, bidiSort: any, ignoreThe: any, ignoreKashida: any, ignoreDiacritics: any, ignoreHe: any, languageID: any): void
        previous(unitObj: any, countObj: any): word.Range
        getTable(): word.Table
        getID(): string
        inRange(range: word.Range): boolean
        isEqual(range: word.Range): boolean
        getFont(): word.Font
        setKana(kana: int): void
        getKana(): int
        getOrientation(): int
        refresh(word: any): void
        checkGrammar(): void
        checkSpelling(customDictionary: any, ignoreUppercase: any, alwaysSuggest: any, customDictionary2: any, customDictionary3: any, customDictionary4: any, customDictionary5: any, customDictionary6: any, customDictionary7: any, customDictionary8: any, customDictionary9: any, customDictionary10: any): void
        getSpellingSuggestions(customDictionary: any, ignoreUppercase: any, mainDictionary: any, suggestionMode: any, customDictionary2: any, customDictionary3: any, customDictionary4: any, customDictionary5: any, customDictionary6: any, customDictionary7: any, customDictionary8: any, customDictionary9: any, customDictionary10: any): word.SpellingSuggestions
        lookupNameProperties(): void
        move(unitObj: any, countObj: any): int
        getFind(): word.Find
        getText(): string
        moveLeft(unitObj: int, countObj: int, extendObj: int): int
        moveLeft(unitObj: any, countObj: any, extendObj: any): int
        moveLeft(unit: int, count: any, extend: any): int
        paste(): void
        setFont(font: word.Font): void
        cut(): void
        getMRange(): any
        getDocument(): word.Document
        moveRight(unitObj: any, countObj: any, extendObj: any): int
        moveRight(unitObj: int, countObj: int, extendObj: int): int
        moveRight(unit: int, count: any, extend: any): int
        insertParagraph(): void
        getFrames(): word.Frames
        getTables(): word.Tables
        setText(text: any): void
        setText(text: string): void
        getStart(): int
        setStart(start: int): void
        setEnd(end: long): void
        setEnd(end: int): void
        getEnd(): int
        select(): void
        insertAfter(text: string): void
        setEmptyObject(isEmpty: boolean): void
        getStoryType(): int
        getPageSetup(): word.PageSetup
        isProtect(): void
        getBookmarks(): word.Bookmarks
        getParagraphs(): word.Paragraphs
        setID(id: string): void
        getCells(): word.Cells
        insertCaption(label: any, title: any, titleAutoText: any, position: any, excludeLabel: any): void
        getSections(): word.Sections
        setBorders(borders: word.Borders): void
        getBorders(): word.Borders
        checkIsNull(): void
        getShading(): word.Shading
        setParagraphFormat(paragraphFormat: word.ParagraphFormat): void
        getParagraphFormat(): word.ParagraphFormat
        computeStatistics(statistic: int): int
        setGrammarChecked(grammarChecked: boolean): void
        getGrammaticalErrors(): word.ProofreadingErrors
        isLanguageDetected(): boolean
        setLanguageDetected(languageDetected: boolean): void
        getReadabilityStatistics(): word.ReadabilityStatistics
        setBold(bold: int): void
        getBold(): int
        pasteAndFormat(type: int): void
        tcscConverter(ywTCSCConverterDirection: int, commonTerms: boolean, useVariants: boolean): void
        setTwoLinesInOne(twoLinesInOne: int): void
        setFitTextWidth(fitTextWidth: double): void
        getFitTextWidth(): double
        getShapeRange(): word.ShapeRange
        pasteSpecial(iconIndex: any, link: any, placement: any, displayAsIcon: any, dataType: any, iconFileName: any, iconLabel: any): void
        setOrientation(orientation: int): void
        convertToTable(separator: any, numRowsObj: any, numColumnsObj: any, initialColumnWidthObj: any, formatObj: any, applyBordersObj: any, applyShadingObj: any, applyFontObj: any, applyColorObj: any, applyHeadingRowsObj: any, applyLastRowObj: any, applyFirstColumnObj: any, applyLastColumnObj: any, autoFitObj: any, autoFitBehaviorObj: any, defaultTableBehaviorObj: any): word.Table
        getNoProofing(): int
        setNoProofing(noProofing: int): void
        getCharacterWidth(): int
        insertParagraphAfter(): void
        insertParagraphBefore(): void
        selectWithoutScroll(): void
        getLanguageIDFarEast(): int
        setLanguageIDFarEast(languageIDFarEast: int): void
        getLanguageIDOther(): int
        setLanguageIDOther(languageIDOther: int): void
        setBoldBi(boldBi: int): void
        getBoldBi(): int
        setEmphasisMark(emphasisMark: int): void
        getEmphasisMark(): int
        setItalic(italic: int): void
        getItalic(): int
        setItalicBi(italicBi: int): void
        getItalicBi(): int
        setUnderline(underline: int): void
        getUnderline(): int
        insertAutoText(): void
        insertSymbol(characterNumber: int, font: any, unicode: any, bias: any): void
        isDisableCharacterSpaceGrid(): boolean
        setDisableCharacterSpaceGrid(disableCharacterSpaceGrid: boolean): void
        getColumns(): word.Columns
        getDuplicate(): word.Range
        getRows(): word.Rows
        getXML(dataOnly: boolean): string
        getCase(): int
        setCase(case1: int): void
        getInlineShapes(): word.InlineShapes
        insertFile(fileName: string, range: any, confirmConversions: any, link: any, attachment: any): void
        getRevisions(): word.Revisions
        getScripts(): office.Scripts
        getSentences(): word.Sentences
        getSmartTags(): word.SmartTags
        getSubdocuments(): word.Subdocuments
        getXMLNodes(): word.XMLNodes
        getWords(): word.Words
        getStyle(): any
        setRange(start: int, end: int): void
        setStyle(style: any): void
        goToNext(what: int): word.Range
        collapse(direction: any): void
        isSpellingChecked(): boolean
        setSpellingChecked(spellingChecked: boolean): void
        getSpellingErrors(): word.ProofreadingErrors
        getEditors(): word.Editors
        goTo(what: any, which: any, count: any, name: any): word.Range
        getLanguageID(): int
        setLanguageID(languageID: int): void
        autoFormat(): void
        getCharacters(): word.Characters
        getComments(): word.Comments
        detectLanguage(): void
        getEndnotes(): word.Endnotes
        getFootnotes(): word.Footnotes
        getFormFields(): word.FormFields
        isGrammarChecked(): boolean
        getHTMLDivisions(): word.HTMLDivisions
        getHyperlinks(): word.Hyperlinks
        getListParagraphs(): word.ListParagraphs
        goToPrevious(what: int): word.Range
        getEndnoteOptions(): word.EndnoteOptions
        getEnhMetaFileBits(): any
        getFootnoteOptions(): word.FootnoteOptions
        getNextStoryRange(): word.Range
        getParagraphStyle(): any
        getPreviousBookmarkID(): int
        getTopLevelTables(): word.Tables
        getTabStops(): word.TabStops
        setTabStops(tabStops: word.TabStops): void
        modifyEnclosure(style: any, symbol: any, enclosedText: any): void
        checkUpOffset(word: any): boolean
        moveEndUntil(Cset: any, count: any): int
        moveEndWhile(Cset: any, count: any): int
        moveStart(unit: any, count: any): int
        moveStartUntil(Cset: any, count: any): int
        moveStartWhile(Cset: any, count: any): int
        moveUntil(Cset: any, count: any): int
        moveWhile(Cset: any, count: any): int
        nextSubdocument(): void
        pasteAppendTable(): void
        pasteExcelTable(linkedToExcel: boolean, wordFormatting: boolean, rtf: boolean): void
        phoneticGuide(text: string, alignment: int, raise: int, fontSize: int, fontName: string): void
        sortAscending(): void
        sortDescending(): void
        subscribeTo(edition: string, format: int): void
        wholeStory(): void
        moveStartInTable(count: int): void
        moveEndInTable(count: int): void
        getTextAsHtml(charsetName: string): string
        getTextAsHtml(): string
        getTextAsMht(): string
        isEndOfRowMark(): boolean
        isEmptyObject(): boolean
        getBookmarkID(): int
        getInformation(type: int): any
        getListFormat(): word.ListFormat
        getStoryLength(): int
        getSynonymInfo(): word.SynonymInfo
        getXMLParentNode(): word.XMLNode
        getFormattedText(): word.Range
        setFormattedText(formattedText: word.Range): void
        isShowAll(): boolean
        setShowAll(showAll: boolean): void
        getTwoLinesInOne(): int
        calculate(): double
        checkSynonyms(): void
        copyAsPicture(): void
        createPublisher(edition: any, containsPICT: boolean, containsRTF: boolean, containsText: boolean): void
        insertBefore(text: string): void
        insertBreak(type: any): void
        insertDatabase(format: any, style: any, linkToSource: any, connection: any, SQLStatement: any, SQLStatement1: any, passwordDocument: any, PasswordTemplate: any, writePasswordDocument: any, writePasswordTemplate: any, dataSource: any, from: any, to: any, includeFields: any): void
        insertDateTime(dateTimeFormat: any, insertAsField: any, insertAsFullWidth: any, dateLanguage: any, calendarType: any): void
        insertFile1(fileName: string, range: any, confirmConversions: any, link: any, attachment: any): boolean
        insertXML(text: string, transform: any): void
        stopScreenUpdating(): void
        startScreenUpdating(): void
        setCharacterWidth(characterWidth: int): void
        isCombineCharacters(): boolean
        setCombineCharacters(combineCharacters: boolean): void
        getHighlightColorIndex(): int
        setHighlightColorIndex(highlightColorIndex: int): void
        getHorizontalInVertical(): int
        setHorizontalInVertical(horizontalInVertical: int): void
        translateToMSText(text: string): string
        getTextRetrievalMode(): word.TextRetrievalMode
        setTextRetrievalMode(textRetrievalMode: word.TextRetrievalMode): void
        convertHangulAndHanja(conversionsMode: any, fastConversion: any, checkHangulEnding: any, enableRecentOrdering: any, customDictionary: any): void
        goToEditableRange(editorID: any): word.Range
        insertCrossReference(referenceType: any, referenceKind: int, referenceItem: any, insertAsHyperlink: any, includePosition: any, separateNumbers: any, separatorString: any): void
        pasteAsNestedTable(): void
        previousSubdocument(): void
        endOf(unit: any, extend: any): int
        doStart(): void
        doEnd(): void
        inStory(range: word.Range): boolean
        moveEnd(unit: int, count: int): int
        moveEnd(unitObj: any, countObj: any): int
        relocate(direction: int): void
        startOf(unit: any, extend: any): int
    }
    export interface ReadabilityStatistic {
        getName(): string
        getValue(): float
    }
    export interface ReadabilityStatistic {
        getName(): string
        getValue(): float
    }
    export interface ReadabilityStatistics {
        item(index: any): word.ReadabilityStatistic
        getCount(): int
    }
    export interface ReadabilityStatistics {
        item(index: any): word.ReadabilityStatistic
        getCount(): int
    }
    export interface RecentFile {
        getName(): string
        delete(): void
        setReadOnly(readOnly: boolean): void
        getPath(): string
        isReadOnly(): boolean
        open(): word.Document
        getIndex(): int
    }
    export interface RecentFile {
        getName(): string
        delete(): void
        setReadOnly(readOnly: boolean): void
        getPath(): string
        isReadOnly(): boolean
        open(): word.Document
        getIndex(): int
    }
    export interface RecentFiles {
        add(document: any, readOnly: boolean): word.RecentFile
        removeAll(): void
        item(index: int): word.RecentFile
        getCount(): int
        getMaximum(): int
        getMRecentFiles(): any
        setMaximum(maximum: int): void
    }
    export interface RecentFiles {
        add(document: any, readOnly: boolean): word.RecentFile
        removeAll(): void
        item(index: int): word.RecentFile
        getCount(): int
        getMaximum(): int
        getMRecentFiles(): any
        setMaximum(maximum: int): void
    }
    export interface Rectangle {
        getWidth(): int
        getLeft(): int
        getTop(): int
        getHeight(): int
        getRange(): word.Range
        getRectangleType(): int
        getLines(): word.Lines
    }
    export interface Rectangle {
        getWidth(): int
        getLeft(): int
        getTop(): int
        getHeight(): int
        getRange(): word.Range
        getRectangleType(): int
        getLines(): word.Lines
    }
    export interface Rectangles {
        item(index: int): word.Rectangle
        getCount(): int
    }
    export interface Rectangles {
        item(index: int): word.Rectangle
        getCount(): int
    }
    export interface Replacement {
        getFont(): word.Font
        getText(): string
        setFont(pFont: word.Font): void
        getFrame(): word.Frame
        setText(pText: string): void
        setParagraphFormat(pParagraphFormat: word.ParagraphFormat): void
        getParagraphFormat(): word.ParagraphFormat
        getHighlight(): int
        clearFormatting(): void
        setHighlight(pHighlight: int): void
        getNoProofing(): int
        setNoProofing(pNoProofing: int): void
        getLanguageIDFarEast(): int
        setLanguageIDFarEast(pLanguageIDFarEast: int): void
        getStyle(): any
        setStyle(pStyle: any): void
        getLanguageID(): int
        setLanguageID(pLanguageID: int): void
    }
    export interface Replacement {
        getFont(): word.Font
        getText(): string
        setFont(pFont: word.Font): void
        getFrame(): word.Frame
        setText(pText: string): void
        setParagraphFormat(pParagraphFormat: word.ParagraphFormat): void
        getParagraphFormat(): word.ParagraphFormat
        getHighlight(): int
        clearFormatting(): void
        setHighlight(pHighlight: int): void
        getNoProofing(): int
        setNoProofing(pNoProofing: int): void
        getLanguageIDFarEast(): int
        setLanguageIDFarEast(pLanguageIDFarEast: int): void
        getStyle(): any
        setStyle(pStyle: any): void
        getLanguageID(): int
        setLanguageID(pLanguageID: int): void
    }
    export interface Reviewer {
        isVisible(): boolean
        setVisible(pVisible: boolean): void
        getUserName(): string
    }
    export interface Reviewer {
        isVisible(): boolean
        setVisible(pVisible: boolean): void
        getUserName(): string
    }
    export interface Reviewers {
        item(index: any): word.Reviewer
        getCount(): int
    }
    export interface Reviewers {
        item(index: any): word.Reviewer
        getCount(): int
    }
    export interface Revision {
        accept(): void
        getType(): int
        getDate(): any
        getIndex(): int
        getRange(): word.Range
        getAuthor(): string
        getStyle(): word.Style
        getRevisionDate(): double
        getFormatDescription(): string
        reject(): void
    }
    export interface Revision {
        accept(): void
        getType(): int
        getDate(): any
        getIndex(): int
        getRange(): word.Range
        getAuthor(): string
        getStyle(): word.Style
        getRevisionDate(): double
        getFormatDescription(): string
        reject(): void
    }
    export interface Revisions {
        item(index: int): word.Revision
        getCount(): int
        setViewType(type: int): void
        acceptAll(): void
        rejectAll(): void
    }
    export interface Revisions {
        item(index: int): word.Revision
        getCount(): int
        setViewType(type: int): void
        acceptAll(): void
        rejectAll(): void
    }
    export interface RevisionsFilter {
        getView(): int
        getReviewers(): word.Reviewers
        setMarkup(type: int): void
        getMarkup(): int
        toggleShowAllReviewers(): void
        setView(type: int): void
    }
    export interface RevisionsFilter {
        getView(): int
        getReviewers(): word.Reviewers
        setMarkup(type: int): void
        getMarkup(): int
        toggleShowAllReviewers(): void
        setView(type: int): void
    }
    export interface RoutingSlip {
        getMessage(): string
        reset(): void
        setProtect(pProtect: int): void
        getRecipients(): string
        getRecipients(index: string): string
        getSubject(): string
        setSubject(pSubject: string): void
        getDelivery(): int
        getProtect(): int
        isReturnWhenDone(): boolean
        getStatus(): int
        isTrackStatus(): boolean
        setDelivery(pDelivery: int): void
        setMessage(pMessage: string): void
        setTrackStatus(pTrackStatus: boolean): void
        addRecipient(recipient: string): void
        setReturnWhenDone(pReturnWhenDone: boolean): void
    }
    export interface RoutingSlip {
        getMessage(): string
        reset(): void
        setProtect(pProtect: int): void
        getRecipients(): string
        getRecipients(index: string): string
        getSubject(): string
        setSubject(pSubject: string): void
        getDelivery(): int
        getProtect(): int
        isReturnWhenDone(): boolean
        getStatus(): int
        isTrackStatus(): boolean
        setDelivery(pDelivery: int): void
        setMessage(pMessage: string): void
        setTrackStatus(pTrackStatus: boolean): void
        addRecipient(recipient: string): void
        setReturnWhenDone(pReturnWhenDone: boolean): void
    }
    export interface Row {
        delete(): void
        getID(): string
        getIndex(): int
        getHeight(): double
        setHeight(pHeight: double): void
        setHeight(rowHeight: float, heightRule: int): void
        getMRange(): any
        setHeightRule(pHeightRule: int): void
        getHeightRule(): int
        getNestingLevel(): int
        getPrevious(): word.Row
        getRange(): word.Range
        select(): void
        setID(pID: string): void
        getNext(): word.Row
        getCells(): word.Cells
        setBorders(pBorders: word.Borders): void
        getMTable(): any
        isFirst(): boolean
        isLast(): boolean
        getBorders(): word.Borders
        checkIsNull(): void
        getShading(): word.Shading
        getAlignment(): int
        setAlignment(pAlignment: int): void
        getLeftIndent(): double
        setLeftIndent(pLeftIndent: double): void
        setLeftIndent(leftIndent: float, rulerStyle: int): void
        getHeadingFormat(): int
        setHeadingFormat(pHeadingFormat: int): void
        convertToText(separator: any, nestedTables: any): word.Range
        getAllowBreakAcrossPages(): int
        setAllowBreakAcrossPages(pAllowBreakAcrossPages: int): void
        getSpaceBetweenColumns(): double
        setSpaceBetweenColumns(pSpaceBetweenColumns: double): void
    }
    export interface Row {
        delete(): void
        getID(): string
        getIndex(): int
        getHeight(): double
        setHeight(pHeight: double): void
        setHeight(rowHeight: float, heightRule: int): void
        getMRange(): any
        setHeightRule(pHeightRule: int): void
        getHeightRule(): int
        getNestingLevel(): int
        getPrevious(): word.Row
        getRange(): word.Range
        select(): void
        setID(pID: string): void
        getNext(): word.Row
        getCells(): word.Cells
        setBorders(pBorders: word.Borders): void
        getMTable(): any
        isFirst(): boolean
        isLast(): boolean
        getBorders(): word.Borders
        checkIsNull(): void
        getShading(): word.Shading
        getAlignment(): int
        setAlignment(pAlignment: int): void
        getLeftIndent(): double
        setLeftIndent(pLeftIndent: double): void
        setLeftIndent(leftIndent: float, rulerStyle: int): void
        getHeadingFormat(): int
        setHeadingFormat(pHeadingFormat: int): void
        convertToText(separator: any, nestedTables: any): word.Range
        getAllowBreakAcrossPages(): int
        setAllowBreakAcrossPages(pAllowBreakAcrossPages: int): void
        getSpaceBetweenColumns(): double
        setSpaceBetweenColumns(pSpaceBetweenColumns: double): void
    }
    export interface Rows {
        add(beforeRowObj: any): word.Row
        getFirst(): word.Row
        delete(): void
        item(index: int): word.Row
        getLast(): word.Row
        getCount(): int
        getHeight(): double
        setHeight(rowHeight: float, heightRule: int): void
        setHeight(pHeight: double): void
        getMRange(): any
        setHeightRule(pHeightRule: int): void
        getHeightRule(): int
        getNestingLevel(): int
        select(): void
        setBorders(pBorders: word.Borders): void
        getMTable(): any
        distributeHeight(): void
        getBorders(): word.Borders
        checkIsNull(): void
        getShading(): word.Shading
        getAlignment(): int
        setAlignment(pAlignment: int): void
        getLeftIndent(): double
        setLeftIndent(pLeftIndent: double): void
        setLeftIndent(leftIndent: float, rulerStyle: int): void
        getRelativeHorizontalPosition(): int
        setRelativeHorizontalPosition(pRelativeHorizontalPosition: int): void
        getRelativeVerticalPosition(): int
        setRelativeVerticalPosition(pRelativeVerticalPosition: int): void
        setHorizontalPosition(pHorizontalPosition: double): void
        getHorizontalPosition(): double
        setVerticalPosition(pVerticalPosition: double): void
        getVerticalPosition(): double
        getHeadingFormat(): int
        setHeadingFormat(pHeadingFormat: int): void
        convertToText(separator: any, nestedTables: any): word.Range
        getAllowOverlap(): int
        setAllowOverlap(pAllowOverlap: int): void
        getDistanceLeft(): double
        setDistanceLeft(pDistanceLeft: double): void
        getDistanceRight(): double
        setDistanceRight(pDistanceRight: double): void
        getDistanceTop(): double
        setDistanceTop(pDistanceTop: double): void
        getParentOfRow(): word.Table
        getAllowBreakAcrossPages(): int
        setAllowBreakAcrossPages(pAllowBreakAcrossPages: int): void
        getSpaceBetweenColumns(): double
        setSpaceBetweenColumns(pSpaceBetweenColumns: double): void
        getDistanceBottom(): double
        setDistanceBottom(pDistanceBottom: double): void
        getTableDirection(): int
        setTableDirection(pTableDirection: int): void
        getWrapAroundText(): int
        setWrapAroundText(pWrapAroundText: int): void
    }
    export interface Rows {
        add(beforeRowObj: any): word.Row
        getFirst(): word.Row
        delete(): void
        item(index: int): word.Row
        getLast(): word.Row
        getCount(): int
        getHeight(): double
        setHeight(rowHeight: float, heightRule: int): void
        setHeight(pHeight: double): void
        getMRange(): any
        setHeightRule(pHeightRule: int): void
        getHeightRule(): int
        getNestingLevel(): int
        select(): void
        setBorders(pBorders: word.Borders): void
        getMTable(): any
        distributeHeight(): void
        getBorders(): word.Borders
        checkIsNull(): void
        getShading(): word.Shading
        getAlignment(): int
        setAlignment(pAlignment: int): void
        getLeftIndent(): double
        setLeftIndent(pLeftIndent: double): void
        setLeftIndent(leftIndent: float, rulerStyle: int): void
        getRelativeHorizontalPosition(): int
        setRelativeHorizontalPosition(pRelativeHorizontalPosition: int): void
        getRelativeVerticalPosition(): int
        setRelativeVerticalPosition(pRelativeVerticalPosition: int): void
        setHorizontalPosition(pHorizontalPosition: double): void
        getHorizontalPosition(): double
        setVerticalPosition(pVerticalPosition: double): void
        getVerticalPosition(): double
        getHeadingFormat(): int
        setHeadingFormat(pHeadingFormat: int): void
        convertToText(separator: any, nestedTables: any): word.Range
        getAllowOverlap(): int
        setAllowOverlap(pAllowOverlap: int): void
        getDistanceLeft(): double
        setDistanceLeft(pDistanceLeft: double): void
        getDistanceRight(): double
        setDistanceRight(pDistanceRight: double): void
        getDistanceTop(): double
        setDistanceTop(pDistanceTop: double): void
        getParentOfRow(): word.Table
        getAllowBreakAcrossPages(): int
        setAllowBreakAcrossPages(pAllowBreakAcrossPages: int): void
        getSpaceBetweenColumns(): double
        setSpaceBetweenColumns(pSpaceBetweenColumns: double): void
        getDistanceBottom(): double
        setDistanceBottom(pDistanceBottom: double): void
        getTableDirection(): int
        setTableDirection(pTableDirection: int): void
        getWrapAroundText(): int
        setWrapAroundText(pWrapAroundText: int): void
    }
    export interface ScopeFolder {
        getName(): string
        getPath(): string
        getScopeFolders(): word.ScopeFolders
        addToSearchFolders(): void
    }
    export interface ScopeFolder {
        getName(): string
        getPath(): string
        getScopeFolders(): word.ScopeFolders
        addToSearchFolders(): void
    }
    export interface ScopeFolders {
        item(index: int): word.ScopeFolder
        getCount(): int
    }
    export interface ScopeFolders {
        item(index: int): word.ScopeFolder
        getCount(): int
    }
    export interface SearchFolders {
        add(folder: word.ScopeFolder): void
        remove(index: int): void
        item(index: int): word.ScopeFolder
        getCount(): int
    }
    export interface SearchFolders {
        add(folder: word.ScopeFolder): void
        remove(index: int): void
        item(index: int): word.ScopeFolder
        getCount(): int
    }
    export interface Section {
        getHeaders(): word.HeadersFooters
        getIndex(): int
        getRange(): word.Range
        getPageSetup(): word.PageSetup
        setBorders(pBorders: word.Borders): void
        getBorders(): word.Borders
        setPageSetup(pPageSetup: word.PageSetup): void
        getFooters(): word.HeadersFooters
        getSection(): any
        isProtectedForForms(): boolean
        setProtectedForForms(pProtectedForForms: boolean): void
    }
    export interface Section {
        getHeaders(): word.HeadersFooters
        getIndex(): int
        getRange(): word.Range
        getPageSetup(): word.PageSetup
        setBorders(pBorders: word.Borders): void
        getBorders(): word.Borders
        setPageSetup(pPageSetup: word.PageSetup): void
        getFooters(): word.HeadersFooters
        getSection(): any
        isProtectedForForms(): boolean
        setProtectedForForms(pProtectedForForms: boolean): void
    }
    export interface Sections {
        add(rangeObj: any, startObj: any): word.Section
        getFirst(): word.Section
        item(index: int): word.Section
        getLast(): word.Section
        getCount(): int
        getPageSetup(): word.PageSetup
    }
    export interface Sections {
        add(rangeObj: any, startObj: any): word.Section
        getFirst(): word.Section
        item(index: int): word.Section
        getLast(): word.Section
        getCount(): int
        getPageSetup(): word.PageSetup
    }
    export interface Selection {
        next(unit: any, count: any): word.Range
        getFields(): word.Fields
        delete(unit: any, count: any): int
        copy(): void
        getType(): int
        expand(unit: any): int
        sort(excludeHeader: any, fieldNumber: any, sortFieldType: any, sortOrder: any, fieldNumber2: any, sortFieldType2: any, sortOrder2: any, fieldNumber3: any, sortFieldType3: any, sortOrder3: any, sortColumn: any, separator: any, caseSensitive: any, bidiSort: any, ignoreThe: any, ignoreKashida: any, ignoreDiacritics: any, ignoreHe: any, languageID: any, subFieldNumber: any, subFieldNumber2: any, subFieldNumber3: any): void
        previous(unit: any, count: any): word.Range
        checkRange(): void
        inRange(range: word.Range): boolean
        isEqual(range: word.Range): boolean
        getFont(): word.Font
        isActive(): boolean
        getOrientation(): int
        move(unit: any, count: any): int
        getFind(): word.Find
        getText(): string
        moveLeft(unitObj: any, count: any, extend: any): int
        moveUp(unit: any, count: any, extend: any): int
        moveDown(unit: any, count: any, extend: any): int
        paste(): void
        setFont(pFont: word.Font): void
        cut(): void
        getDocument(): word.Document
        moveRight(unit: int, count: int, extend: int): int
        moveRight(unit: any, count: any, extend: any): int
        moveRight(unit: int, count: any, extend: any): int
        insertParagraph(): void
        getFrames(): word.Frames
        getTables(): word.Tables
        setText(pText: string): void
        getStart(): int
        setStart(pStart: int): void
        setEnd(pEnd: int): void
        getEnd(): int
        getRange(): word.Range
        select(): void
        insertAfter(text: string): void
        getStoryType(): int
        getPageSetup(): word.PageSetup
        isProtect(): void
        getBookmarks(): word.Bookmarks
        getParagraphs(): word.Paragraphs
        getCells(): word.Cells
        insertCaption(label: any, title: any, titleAutoText: any, position: any, excludeLabel: any): void
        getSections(): word.Sections
        setBorders(pBorders: word.Borders): void
        getBorders(): word.Borders
        selectCell(): void
        getShading(): word.Shading
        setParagraphFormat(pParagraphFormat: word.ParagraphFormat): void
        getParagraphFormat(): word.ParagraphFormat
        selectColumn(): void
        isLanguageDetected(): boolean
        setLanguageDetected(pLanguageDetected: boolean): void
        shrink(): void
        pasteAndFormat(type: int): void
        copyFormat(): void
        pasteFormat(): void
        setFitTextWidth(pFitTextWidth: double): void
        getFitTextWidth(): double
        setPageSetup(pPageSetup: word.PageSetup): void
        getShapeRange(): word.ShapeRange
        pasteSpecial(iconIndex: any, link: any, placement: any, displayAsIcon: any, dataType: any, iconFileName: any, iconLabel: any): void
        setOrientation(pOrientation: int): void
        convertToTable(separator: any, numRows: any, numColumns: any, initialColumnWidth: any, format: any, applyBorders: any, applyShading: any, applyFont: any, applyColor: any, applyHeadingRows: any, applyLastRow: any, applyFirstColumn: any, applyLastColumn: any, autoFit: any, autoFitBehavior: any, defaultTableBehavior: any): word.Table
        clearFormatting(): void
        getNoProofing(): int
        setNoProofing(pNoProofing: int): void
        insertParagraphAfter(): void
        insertParagraphBefore(): void
        getLanguageIDFarEast(): int
        setLanguageIDFarEast(pLanguageIDFarEast: int): void
        getLanguageIDOther(): int
        setLanguageIDOther(pLanguageIDOther: int): void
        insertSymbol(characterNumber: int, font: any, unicode: any, bias: any): void
        getColumns(): word.Columns
        getRows(): word.Rows
        getXML(dataOnly: boolean): string
        getInlineShapes(): word.InlineShapes
        insertFile(fileName: string, range: any, confirmConversions: any, link: any, attachment: any): void
        getSentences(): word.Sentences
        getSmartTags(): word.SmartTags
        getXMLNodes(): word.XMLNodes
        getWords(): word.Words
        getStyle(): word.Style
        setRange(start: int, end: int): void
        setStyle(pStyle: any): void
        goToNext(what: int): word.Range
        collapse(direction: any): void
        getEditors(): word.Editors
        goTo(what: any, which: any, count: any, name: any): word.Range
        getLanguageID(): int
        setLanguageID(pLanguageID: int): void
        getCharacters(): word.Characters
        getComments(): word.Comments
        detectLanguage(): void
        getEndnotes(): word.Endnotes
        getFootnotes(): word.Footnotes
        getFormFields(): word.FormFields
        getHTMLDivisions(): word.HTMLDivisions
        getHyperlinks(): word.Hyperlinks
        goToPrevious(what: int): word.Range
        getEndnoteOptions(): word.EndnoteOptions
        getEnhMetaFileBits(): any
        getFootnoteOptions(): word.FootnoteOptions
        getPreviousBookmarkID(): int
        getTopLevelTables(): word.Tables
        checkUpOffset(word: any): boolean
        moveEndUntil(cset: any, count: any): int
        moveEndWhile(cset: any, count: any): int
        moveStart(unit: any, count: any): int
        moveStartUntil(cset: any, count: any): int
        moveStartWhile(cset: any, count: any): int
        moveUntil(cset: any, count: any): int
        moveWhile(cset: any, count: any): int
        nextSubdocument(): void
        pasteAppendTable(): void
        pasteExcelTable(linkedToExcel: boolean, wordFormatting: boolean, rtf: boolean): void
        sortAscending(): void
        sortDescending(): void
        wholeStory(): void
        selectRow(): void
        isExtendMode(): boolean
        setExtendMode(pExtendMode: boolean): void
        getHeaderFooter(): word.HeaderFooter
        isIPAtEndOfLine(): boolean
        IsEndOfRowMark(): boolean
        isStartIsActive(): boolean
        setStartIsActive(pStartIsActive: boolean): void
        dataExChanged(): void
        createTextbox(): void
        escapeKey(): void
        insertCells(shiftCells: any): void
        insertColumns(): void
        insertColumns(isleft: boolean): void
        insertFormula(formula: any, numberFormat: any): void
        insertNewPage(): void
        insertRows(obj: any, type: boolean): any
        insertRows(numRows: any): void
        insertRowsAbove(obj: any): void
        insertRowsBelow(obj: any): void
        italicRun(): void
        nextField(): word.Field
        nextRevision(wrap: any): word.Revision
        previousField(): word.Field
        previousRevision(wrap: any): word.Revision
        splitTable(): void
        typeBackspace(): void
        typeParagraph(): void
        clearParagraphAllFormatting(): void
        clearParagraphDirectFormatting(): void
        shrinkDiscontiguousSelection(): void
        getChildShapeRange(): word.ShapeRange
        isColumnSelectMode(): boolean
        setColumnSelectMode(pColumnSelectMode: boolean): void
        hasChildShapeRange(): boolean
        clearParagraphStyle(): void
        createAutoTextEntry(name: string, styleName: string): word.AutoTextEntry
        insertColumnsRight(): void
        insertStyleSeparator(): void
        selectCurrentAlignment(): void
        selectCurrentColor(): void
        selectCurrentFont(): void
        selectCurrentIndent(): void
        selectCurrentSpacing(): void
        selectCurrentTabs(): void
        getBookmarkID(): int
        getInformation(type: int): any
        getStoryLength(): int
        getXMLParentNode(): word.XMLNode
        getFormattedText(): word.Range
        setFormattedText(pFormattedText: word.Range): void
        calculate(): double
        copyAsPicture(): void
        insertBefore(text: string): void
        insertBreak(type: any): void
        insertDateTime(dateTimeFormat: any, insertAsField: any, insertAsFullWidth: any, dateLanguage: any, calendarType: any): void
        insertFile1(fileName: string, range: any, confirmConversions: any, link: any, attachment: any): boolean
        insertXML(xml: string, transform: any): void
        toggleCharacterCode(): void
        goToEditableRange(): word.Range
        goToEditableRange(editorID: any): word.Range
        insertCrossReference(referenceType: any, referenceKind: int, referenceItem: any, insertAsHyperlink: any, includePosition: any, separateNumbers: any, separatorString: any): void
        pasteAsNestedTable(): void
        previousSubdocument(): void
        endOf(unit: any, extend: any): int
        extend(character: any): void
        doEnd(select: boolean): void
        doEnd(): void
        inStory(range: word.Range): boolean
        moveEnd(unitObj: any, countObj: any): int
        startOf(unit: any, extend: any): int
        getFlags(): int
        setFlags(pFlags: int): void
        boldRun(): void
        endKey(unit: any, extend: any): int
        homeKey(unit: any, extend: any): int
        ltrPara(): void
        ltrRun(): void
        rtlPara(): void
        rtlRun(): void
        typeText(text: string): void
    }
    export interface Selection {
        next(unit: any, count: any): word.Range
        getFields(): word.Fields
        delete(unit: any, count: any): int
        copy(): void
        getType(): int
        expand(unit: any): int
        sort(excludeHeader: any, fieldNumber: any, sortFieldType: any, sortOrder: any, fieldNumber2: any, sortFieldType2: any, sortOrder2: any, fieldNumber3: any, sortFieldType3: any, sortOrder3: any, sortColumn: any, separator: any, caseSensitive: any, bidiSort: any, ignoreThe: any, ignoreKashida: any, ignoreDiacritics: any, ignoreHe: any, languageID: any, subFieldNumber: any, subFieldNumber2: any, subFieldNumber3: any): void
        previous(unit: any, count: any): word.Range
        checkRange(): void
        inRange(range: word.Range): boolean
        isEqual(range: word.Range): boolean
        getFont(): word.Font
        isActive(): boolean
        getOrientation(): int
        move(unit: any, count: any): int
        getFind(): word.Find
        getText(): string
        moveLeft(unitObj: any, count: any, extend: any): int
        moveUp(unit: any, count: any, extend: any): int
        moveDown(unit: any, count: any, extend: any): int
        paste(): void
        setFont(pFont: word.Font): void
        cut(): void
        getDocument(): word.Document
        moveRight(unit: int, count: int, extend: int): int
        moveRight(unit: any, count: any, extend: any): int
        moveRight(unit: int, count: any, extend: any): int
        insertParagraph(): void
        getFrames(): word.Frames
        getTables(): word.Tables
        setText(pText: string): void
        getStart(): int
        setStart(pStart: int): void
        setEnd(pEnd: int): void
        getEnd(): int
        getRange(): word.Range
        select(): void
        insertAfter(text: string): void
        getStoryType(): int
        getPageSetup(): word.PageSetup
        isProtect(): void
        getBookmarks(): word.Bookmarks
        getParagraphs(): word.Paragraphs
        getCells(): word.Cells
        insertCaption(label: any, title: any, titleAutoText: any, position: any, excludeLabel: any): void
        getSections(): word.Sections
        setBorders(pBorders: word.Borders): void
        getBorders(): word.Borders
        selectCell(): void
        getShading(): word.Shading
        setParagraphFormat(pParagraphFormat: word.ParagraphFormat): void
        getParagraphFormat(): word.ParagraphFormat
        selectColumn(): void
        isLanguageDetected(): boolean
        setLanguageDetected(pLanguageDetected: boolean): void
        shrink(): void
        pasteAndFormat(type: int): void
        copyFormat(): void
        pasteFormat(): void
        setFitTextWidth(pFitTextWidth: double): void
        getFitTextWidth(): double
        setPageSetup(pPageSetup: word.PageSetup): void
        getShapeRange(): word.ShapeRange
        pasteSpecial(iconIndex: any, link: any, placement: any, displayAsIcon: any, dataType: any, iconFileName: any, iconLabel: any): void
        setOrientation(pOrientation: int): void
        convertToTable(separator: any, numRows: any, numColumns: any, initialColumnWidth: any, format: any, applyBorders: any, applyShading: any, applyFont: any, applyColor: any, applyHeadingRows: any, applyLastRow: any, applyFirstColumn: any, applyLastColumn: any, autoFit: any, autoFitBehavior: any, defaultTableBehavior: any): word.Table
        clearFormatting(): void
        getNoProofing(): int
        setNoProofing(pNoProofing: int): void
        insertParagraphAfter(): void
        insertParagraphBefore(): void
        getLanguageIDFarEast(): int
        setLanguageIDFarEast(pLanguageIDFarEast: int): void
        getLanguageIDOther(): int
        setLanguageIDOther(pLanguageIDOther: int): void
        insertSymbol(characterNumber: int, font: any, unicode: any, bias: any): void
        getColumns(): word.Columns
        getRows(): word.Rows
        getXML(dataOnly: boolean): string
        getInlineShapes(): word.InlineShapes
        insertFile(fileName: string, range: any, confirmConversions: any, link: any, attachment: any): void
        getSentences(): word.Sentences
        getSmartTags(): word.SmartTags
        getXMLNodes(): word.XMLNodes
        getWords(): word.Words
        getStyle(): word.Style
        setRange(start: int, end: int): void
        setStyle(pStyle: any): void
        goToNext(what: int): word.Range
        collapse(direction: any): void
        getEditors(): word.Editors
        goTo(what: any, which: any, count: any, name: any): word.Range
        getLanguageID(): int
        setLanguageID(pLanguageID: int): void
        getCharacters(): word.Characters
        getComments(): word.Comments
        detectLanguage(): void
        getEndnotes(): word.Endnotes
        getFootnotes(): word.Footnotes
        getFormFields(): word.FormFields
        getHTMLDivisions(): word.HTMLDivisions
        getHyperlinks(): word.Hyperlinks
        goToPrevious(what: int): word.Range
        getEndnoteOptions(): word.EndnoteOptions
        getEnhMetaFileBits(): any
        getFootnoteOptions(): word.FootnoteOptions
        getPreviousBookmarkID(): int
        getTopLevelTables(): word.Tables
        checkUpOffset(word: any): boolean
        moveEndUntil(cset: any, count: any): int
        moveEndWhile(cset: any, count: any): int
        moveStart(unit: any, count: any): int
        moveStartUntil(cset: any, count: any): int
        moveStartWhile(cset: any, count: any): int
        moveUntil(cset: any, count: any): int
        moveWhile(cset: any, count: any): int
        nextSubdocument(): void
        pasteAppendTable(): void
        pasteExcelTable(linkedToExcel: boolean, wordFormatting: boolean, rtf: boolean): void
        sortAscending(): void
        sortDescending(): void
        wholeStory(): void
        selectRow(): void
        isExtendMode(): boolean
        setExtendMode(pExtendMode: boolean): void
        getHeaderFooter(): word.HeaderFooter
        isIPAtEndOfLine(): boolean
        IsEndOfRowMark(): boolean
        isStartIsActive(): boolean
        setStartIsActive(pStartIsActive: boolean): void
        dataExChanged(): void
        createTextbox(): void
        escapeKey(): void
        insertCells(shiftCells: any): void
        insertColumns(): void
        insertColumns(isleft: boolean): void
        insertFormula(formula: any, numberFormat: any): void
        insertNewPage(): void
        insertRows(obj: any, type: boolean): any
        insertRows(numRows: any): void
        insertRowsAbove(obj: any): void
        insertRowsBelow(obj: any): void
        italicRun(): void
        nextField(): word.Field
        nextRevision(wrap: any): word.Revision
        previousField(): word.Field
        previousRevision(wrap: any): word.Revision
        splitTable(): void
        typeBackspace(): void
        typeParagraph(): void
        clearParagraphAllFormatting(): void
        clearParagraphDirectFormatting(): void
        shrinkDiscontiguousSelection(): void
        getChildShapeRange(): word.ShapeRange
        isColumnSelectMode(): boolean
        setColumnSelectMode(pColumnSelectMode: boolean): void
        hasChildShapeRange(): boolean
        clearParagraphStyle(): void
        createAutoTextEntry(name: string, styleName: string): word.AutoTextEntry
        insertColumnsRight(): void
        insertStyleSeparator(): void
        selectCurrentAlignment(): void
        selectCurrentColor(): void
        selectCurrentFont(): void
        selectCurrentIndent(): void
        selectCurrentSpacing(): void
        selectCurrentTabs(): void
        getBookmarkID(): int
        getInformation(type: int): any
        getStoryLength(): int
        getXMLParentNode(): word.XMLNode
        getFormattedText(): word.Range
        setFormattedText(pFormattedText: word.Range): void
        calculate(): double
        copyAsPicture(): void
        insertBefore(text: string): void
        insertBreak(type: any): void
        insertDateTime(dateTimeFormat: any, insertAsField: any, insertAsFullWidth: any, dateLanguage: any, calendarType: any): void
        insertFile1(fileName: string, range: any, confirmConversions: any, link: any, attachment: any): boolean
        insertXML(xml: string, transform: any): void
        toggleCharacterCode(): void
        goToEditableRange(): word.Range
        goToEditableRange(editorID: any): word.Range
        insertCrossReference(referenceType: any, referenceKind: int, referenceItem: any, insertAsHyperlink: any, includePosition: any, separateNumbers: any, separatorString: any): void
        pasteAsNestedTable(): void
        previousSubdocument(): void
        endOf(unit: any, extend: any): int
        extend(character: any): void
        doEnd(select: boolean): void
        doEnd(): void
        inStory(range: word.Range): boolean
        moveEnd(unitObj: any, countObj: any): int
        startOf(unit: any, extend: any): int
        getFlags(): int
        setFlags(pFlags: int): void
        boldRun(): void
        endKey(unit: any, extend: any): int
        homeKey(unit: any, extend: any): int
        ltrPara(): void
        ltrRun(): void
        rtlPara(): void
        rtlRun(): void
        typeText(text: string): void
    }
    export interface Sentences {
        getFirst(): word.Range
        item(index: int): word.Range
        getLast(): word.Range
        getCount(): int
    }
    export interface Sentences {
        getFirst(): word.Range
        item(index: int): word.Range
        getLast(): word.Range
        getCount(): int
    }
    export interface Shading {
        apply(): void
        setType(type: int): void
        getTexture(): int
        setTexture(pTexture: int): void
        getBackgroundPatternColor(): int
        getBackgroundPatternColorIndex(): int
        getForegroundPatternColor(): int
        getForegroundPatternColorIndex(): int
        setBackgroundPatternColor(pBackgroundPatternColor: int): void
        setBackgroundPatternColorIndex(pBackgroundPatternColorIndex: int): void
        setForegroundPatternColor(pForegroundPatternColor: int): void
        setForegroundPatternColorIndex(pForegroundPatternColorIndex: int): void
    }
    export interface Shading {
        apply(): void
        setType(type: int): void
        getTexture(): int
        setTexture(pTexture: int): void
        getBackgroundPatternColor(): int
        getBackgroundPatternColorIndex(): int
        getForegroundPatternColor(): int
        getForegroundPatternColorIndex(): int
        setBackgroundPatternColor(pBackgroundPatternColor: int): void
        setBackgroundPatternColorIndex(pBackgroundPatternColorIndex: int): void
        setForegroundPatternColor(pForegroundPatternColor: int): void
        setForegroundPatternColorIndex(pForegroundPatternColorIndex: int): void
    }
    export interface ShadowFormat {
        getType(): int
        getSize(): double
        setSize(size: double): void
        setType(pType: int): void
        setVisible(visible: int): void
        getForeColor(): word.ColorFormat
        getTransparency(): double
        setTransparency(pTransparency: double): void
        getVisible(): int
        setForeColor(pForeColor: word.ColorFormat): void
        getObscured(): int
        setObscured(pObscured: int): void
        getOffsetX(): double
        setOffsetX(pOffsetX: double): void
        getOffsetY(): double
        setOffsetY(pOffsetY: double): void
        incrementOffsetX(increment: double): void
        incrementOffsetY(increment: double): void
        getBlur(): double
        setBlur(blur: double): void
    }
    export interface ShadowFormat {
        getType(): int
        getSize(): double
        setSize(size: double): void
        setType(pType: int): void
        setVisible(visible: int): void
        getForeColor(): word.ColorFormat
        getTransparency(): double
        setTransparency(pTransparency: double): void
        getVisible(): int
        setForeColor(pForeColor: word.ColorFormat): void
        getObscured(): int
        setObscured(pObscured: int): void
        getOffsetX(): double
        setOffsetX(pOffsetX: double): void
        getOffsetY(): double
        setOffsetY(pOffsetY: double): void
        incrementOffsetX(increment: double): void
        incrementOffsetY(increment: double): void
        getBlur(): double
        setBlur(blur: double): void
    }
    export interface Shape {
        getName(): string
        apply(): void
        delete(): void
        setName(pName: string): void
        getType(): int
        flip(flipCmd: int): void
        getID(): int
        duplicate(): word.Shape
        getScript(): office.Script
        getShape(): any
        setTitle(text: string): void
        getTitle(): string
        getWidth(): float
        getLeft(): double
        setLeft(pLeft: double): void
        setTop(pTop: double): void
        getTop(): double
        setWidth(pWidth: double): void
        activate(): void
        getHeight(): double
        setHeight(pHeight: double): void
        setVisible(pVisible: int): void
        saveAs(savePath: string): void
        select(replace: any): void
        getFill(): word.FillFormat
        getLine(): word.LineFormat
        getOLEFormat(): word.OLEFormat
        getHyperlink(): word.Hyperlink
        getVisible(): int
        getShadow(): word.ShadowFormat
        setLockAnchor(pLockAnchor: int): void
        getTextFrame(): word.TextFrame
        getRelativeHorizontalPosition(): int
        setRelativeHorizontalPosition(pRelativeHorizontalPosition: int): void
        getRelativeVerticalPosition(): int
        setRelativeVerticalPosition(pRelativeVerticalPosition: int): void
        getPictureFormat(): word.PictureFormat
        getTextEffect(): word.TextEffectFormat
        getGroupItems(): word.GroupShapes
        setAlternativeText(pAlternativeText: string): void
        getAlternativeText(): string
        setLockAspectRatio(pLockAspectRatio: int): void
        getLockAspectRatio(): int
        getLinkFormat(): word.LinkFormat
        getDiagram(): word.Diagram
        ungroup(): word.ShapeRange
        getNodes(): word.ShapeNodes
        setShapesDefaultProperties(): void
        getAdjustments(): word.Adjustments
        getAnchor(): word.Range
        getAutoShapeType(): int
        getCallout(): word.CalloutFormat
        getCanvasItems(): word.CanvasShapes
        getDiagramNode(): word.DiagramNode
        getHasDiagram(): int
        getLayoutInCell(): int
        getLockAnchor(): int
        getParentGroup(): word.Shape
        getRotation(): double
        getThreeD(): word.ThreeDFormat
        isFreeform(shape: any): boolean
        getVerticalFlip(): int
        getVertices(): any
        getWrapFormat(): word.WrapFormat
        setAutoShapeType(pAutoShapeType: int): void
        setLayoutInCell(pLayoutInCell: int): void
        setRotation(pRotation: double): void
        canvasCropBottom(increment: double): void
        canvasCropLeft(increment: double): void
        canvasCropRight(increment: double): void
        canvasCropTop(increment: double): void
        convertToFrame(): word.Frame
        incrementLeft(increment: double): void
        incrementTop(increment: double): void
        scaleHeight(factor: double, relativeToOriginalSize: int, scale: int): void
        scaleWidth(factor: double, relativeToOriginalSize: int, scale: int): void
        setOthersMarkID(markID: string): void
        getOthersMarkID(): string
        getHasDiagramNode(): int
        getHorizontalFlip(): int
        getZOrderPosition(): int
        convertToInlineShape(): word.InlineShape
        incrementRotation(increment: double): void
        setOthersPressValues(pressValues: number[]): void
        getOthersPressValue(): number[]
        setOthersUserName(name: string): void
        getOthersUserName(): string
        setOthersCreateTime(time: string): void
        getOthersCreateTime(): string
        getChild(): int
        isLine(shape: any): boolean
        pickUp(): void
        zOrder(zOrderCmd: int): void
        hasChart(): int
        isIsf(): boolean
    }
    export interface Shape {
        getName(): string
        apply(): void
        delete(): void
        setName(pName: string): void
        getType(): int
        flip(flipCmd: int): void
        getID(): int
        duplicate(): word.Shape
        getScript(): office.Script
        getShape(): any
        setTitle(text: string): void
        getTitle(): string
        getWidth(): float
        getLeft(): double
        setLeft(pLeft: double): void
        setTop(pTop: double): void
        getTop(): double
        setWidth(pWidth: double): void
        activate(): void
        getHeight(): double
        setHeight(pHeight: double): void
        setVisible(pVisible: int): void
        saveAs(savePath: string): void
        select(replace: any): void
        getFill(): word.FillFormat
        getLine(): word.LineFormat
        getOLEFormat(): word.OLEFormat
        getHyperlink(): word.Hyperlink
        getVisible(): int
        getShadow(): word.ShadowFormat
        setLockAnchor(pLockAnchor: int): void
        getTextFrame(): word.TextFrame
        getRelativeHorizontalPosition(): int
        setRelativeHorizontalPosition(pRelativeHorizontalPosition: int): void
        getRelativeVerticalPosition(): int
        setRelativeVerticalPosition(pRelativeVerticalPosition: int): void
        getPictureFormat(): word.PictureFormat
        getTextEffect(): word.TextEffectFormat
        getGroupItems(): word.GroupShapes
        setAlternativeText(pAlternativeText: string): void
        getAlternativeText(): string
        setLockAspectRatio(pLockAspectRatio: int): void
        getLockAspectRatio(): int
        getLinkFormat(): word.LinkFormat
        getDiagram(): word.Diagram
        ungroup(): word.ShapeRange
        getNodes(): word.ShapeNodes
        setShapesDefaultProperties(): void
        getAdjustments(): word.Adjustments
        getAnchor(): word.Range
        getAutoShapeType(): int
        getCallout(): word.CalloutFormat
        getCanvasItems(): word.CanvasShapes
        getDiagramNode(): word.DiagramNode
        getHasDiagram(): int
        getLayoutInCell(): int
        getLockAnchor(): int
        getParentGroup(): word.Shape
        getRotation(): double
        getThreeD(): word.ThreeDFormat
        isFreeform(shape: any): boolean
        getVerticalFlip(): int
        getVertices(): any
        getWrapFormat(): word.WrapFormat
        setAutoShapeType(pAutoShapeType: int): void
        setLayoutInCell(pLayoutInCell: int): void
        setRotation(pRotation: double): void
        canvasCropBottom(increment: double): void
        canvasCropLeft(increment: double): void
        canvasCropRight(increment: double): void
        canvasCropTop(increment: double): void
        convertToFrame(): word.Frame
        incrementLeft(increment: double): void
        incrementTop(increment: double): void
        scaleHeight(factor: double, relativeToOriginalSize: int, scale: int): void
        scaleWidth(factor: double, relativeToOriginalSize: int, scale: int): void
        setOthersMarkID(markID: string): void
        getOthersMarkID(): string
        getHasDiagramNode(): int
        getHorizontalFlip(): int
        getZOrderPosition(): int
        convertToInlineShape(): word.InlineShape
        incrementRotation(increment: double): void
        setOthersPressValues(pressValues: number[]): void
        getOthersPressValue(): number[]
        setOthersUserName(name: string): void
        getOthersUserName(): string
        setOthersCreateTime(time: string): void
        getOthersCreateTime(): string
        getChild(): int
        isLine(shape: any): boolean
        pickUp(): void
        zOrder(zOrderCmd: int): void
        hasChart(): int
        isIsf(): boolean
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
        item(index: any): word.ShapeNode
        getCount(): int
        setPosition(index: int, x1: double, x2: double): void
        setEditingType(index: int, editingType: int): void
        setSegmentType(index: int, segmentType: int): void
    }
    export interface ShapeNodes {
        delete(index: int): void
        insert(index: int, segmentType: int, editingType: int, x1: double, y1: double, x2: double, y2: double, x3: double, y3: double): void
        item(index: any): word.ShapeNode
        getCount(): int
        setPosition(index: int, x1: double, x2: double): void
        setEditingType(index: int, editingType: int): void
        setSegmentType(index: int, segmentType: int): void
    }
    export interface ShapeRange {
        group(): word.Shape
        getName(): string
        apply(): void
        delete(): void
        setName(pName: string): void
        getType(): int
        flip(flipCmd: int): void
        getID(): int
        duplicate(): word.ShapeRange
        item(index: any): word.Shape
        getCount(): int
        setTitle(text: string): void
        getTitle(): string
        getWidth(): double
        getLeft(): double
        setLeft(pLeft: double): void
        setTop(pTop: double): void
        getTop(): double
        setWidth(pWidth: double): void
        activate(): void
        getHeight(): double
        setHeight(pHeight: double): void
        setVisible(pVisible: int): void
        select(replace: any): void
        align(align: int, lelativeTo: int): void
        getFill(): word.FillFormat
        getLine(): word.LineFormat
        getHyperlink(): word.Hyperlink
        getShapeRange(): any
        getVisible(): int
        getShadow(): word.ShadowFormat
        setLockAnchor(pLockAnchor: int): void
        getMShape(): any
        getTextFrame(): word.TextFrame
        getRelativeHorizontalPosition(): int
        setRelativeHorizontalPosition(pRelativeHorizontalPosition: int): void
        getRelativeVerticalPosition(): int
        setRelativeVerticalPosition(pRelativeVerticalPosition: int): void
        getPictureFormat(): word.PictureFormat
        getTextEffect(): word.TextEffectFormat
        getGroupItems(): word.GroupShapes
        setAlternativeText(pAlternativeText: string): void
        getAlternativeText(): string
        setLockAspectRatio(pLockAspectRatio: int): void
        getLockAspectRatio(): int
        getDiagram(): word.Diagram
        ungroup(): word.ShapeRange
        getNodes(): word.ShapeNodes
        setShapesDefaultProperties(): void
        regroup(): word.Shape
        getAdjustments(): word.Adjustments
        getAnchor(): word.Range
        getAutoShapeType(): int
        getCallout(): word.CalloutFormat
        getCanvasItems(): word.CanvasShapes
        getDiagramNode(): word.DiagramNode
        getHasDiagram(): int
        getLayoutInCell(): int
        getLockAnchor(): int
        getParentGroup(): word.Shape
        getRotation(): double
        getThreeD(): word.ThreeDFormat
        getVerticalFlip(): int
        getVertices(): any
        getWrapFormat(): word.WrapFormat
        setAutoShapeType(pAutoShapeType: int): void
        setLayoutInCell(pLayoutInCell: int): void
        setRotation(pRotation: double): void
        canvasCropBottom(increment: double): void
        canvasCropLeft(increment: double): void
        canvasCropRight(increment: double): void
        canvasCropTop(increment: double): void
        convertToFrame(): word.Frame
        incrementLeft(increment: double): void
        incrementTop(increment: double): void
        scaleHeight(factor: double, relativeToOriginalSize: int, scale: int): void
        scaleWidth(factor: double, relativeToOriginalSize: int, scale: int): void
        distribute(distribute: int, relativeTo: int): void
        getHasDiagramNode(): int
        getHorizontalFlip(): int
        getZOrderPosition(): int
        convertToInlineShape(): word.InlineShape
        incrementRotation(increment: double): void
        getChild(): int
        pickUp(): void
        zOrder(zOrderCmd: int): void
    }
    export interface ShapeRange {
        group(): word.Shape
        getName(): string
        apply(): void
        delete(): void
        setName(pName: string): void
        getType(): int
        flip(flipCmd: int): void
        getID(): int
        duplicate(): word.ShapeRange
        item(index: any): word.Shape
        getCount(): int
        setTitle(text: string): void
        getTitle(): string
        getWidth(): double
        getLeft(): double
        setLeft(pLeft: double): void
        setTop(pTop: double): void
        getTop(): double
        setWidth(pWidth: double): void
        activate(): void
        getHeight(): double
        setHeight(pHeight: double): void
        setVisible(pVisible: int): void
        select(replace: any): void
        align(align: int, lelativeTo: int): void
        getFill(): word.FillFormat
        getLine(): word.LineFormat
        getHyperlink(): word.Hyperlink
        getShapeRange(): any
        getVisible(): int
        getShadow(): word.ShadowFormat
        setLockAnchor(pLockAnchor: int): void
        getMShape(): any
        getTextFrame(): word.TextFrame
        getRelativeHorizontalPosition(): int
        setRelativeHorizontalPosition(pRelativeHorizontalPosition: int): void
        getRelativeVerticalPosition(): int
        setRelativeVerticalPosition(pRelativeVerticalPosition: int): void
        getPictureFormat(): word.PictureFormat
        getTextEffect(): word.TextEffectFormat
        getGroupItems(): word.GroupShapes
        setAlternativeText(pAlternativeText: string): void
        getAlternativeText(): string
        setLockAspectRatio(pLockAspectRatio: int): void
        getLockAspectRatio(): int
        getDiagram(): word.Diagram
        ungroup(): word.ShapeRange
        getNodes(): word.ShapeNodes
        setShapesDefaultProperties(): void
        regroup(): word.Shape
        getAdjustments(): word.Adjustments
        getAnchor(): word.Range
        getAutoShapeType(): int
        getCallout(): word.CalloutFormat
        getCanvasItems(): word.CanvasShapes
        getDiagramNode(): word.DiagramNode
        getHasDiagram(): int
        getLayoutInCell(): int
        getLockAnchor(): int
        getParentGroup(): word.Shape
        getRotation(): double
        getThreeD(): word.ThreeDFormat
        getVerticalFlip(): int
        getVertices(): any
        getWrapFormat(): word.WrapFormat
        setAutoShapeType(pAutoShapeType: int): void
        setLayoutInCell(pLayoutInCell: int): void
        setRotation(pRotation: double): void
        canvasCropBottom(increment: double): void
        canvasCropLeft(increment: double): void
        canvasCropRight(increment: double): void
        canvasCropTop(increment: double): void
        convertToFrame(): word.Frame
        incrementLeft(increment: double): void
        incrementTop(increment: double): void
        scaleHeight(factor: double, relativeToOriginalSize: int, scale: int): void
        scaleWidth(factor: double, relativeToOriginalSize: int, scale: int): void
        distribute(distribute: int, relativeTo: int): void
        getHasDiagramNode(): int
        getHorizontalFlip(): int
        getZOrderPosition(): int
        convertToInlineShape(): word.InlineShape
        incrementRotation(increment: double): void
        getChild(): int
        pickUp(): void
        zOrder(zOrderCmd: int): void
    }
    export interface Shapes {
        range(index: any): word.ShapeRange
        item(index: any): word.Shape
        getShape(shapeName: string): word.Shape
        getCount(): int
        addCallout(type: int, left: double, top: double, width: double, height: double, anchor: any): word.Shape
        addPicture(fileName: string, linkToFileObj: any, saveWithDocumentObj: any, leftObj: any, topObj: any, widthObj: any, heightObj: any, anchorObj: any): word.Shape
        addPolyline(safeArrayOfPoints: any, anchor: any): word.Shape
        addTextbox(orientation: int, left: double, top: double, width: double, height: double, anchor: any): word.Shape
        addTextEffect(presetTextEffect: int, text: string, fontName: string, fontSize: double, fontBold: int, fontItalic: int, left: double, top: double, anchor: any): word.Shape
        buildFreeform(editingType: int, x1: double, y1: double): word.FreeformBuilder
        selectAll(): void
        addCurve(safeArrayOfPoints: any, anchor: any): word.Shape
        addLabel(orientation: int, left: double, top: double, width: double, height: double, anchor: any): word.Shape
        addLine(beginX: double, beginY: double, endX: double, endY: double, anchor: any): word.Shape
        addShape(type: int, left: double, top: double, width: double, height: double, anchor: any): word.Shape
        addOLEControl(classType: any, left: any, top: any, width: any, height: any, anchor: any): word.Shape
        addOLEObject(classType: any, fileName: any, linkToFile: any, displayAsIcon: any, iconFileName: any, iconIndex: any, iconLabel: any, left: any, top: any, width: any, height: any, anchor: any): word.Shape
        getMShapes(): any
        checkDoc(): void
        getAllShapes(): word.Shape[]
        addCanvas(left: double, top: double, width: double, height: double, anchor: any): word.Shape
        addDiagram(type: int, left: float, top: float, width: float, height: float, anchor: any): word.Shape
        checkPicture(): void
    }
    export interface Shapes {
        range(index: any): word.ShapeRange
        item(index: any): word.Shape
        getShape(shapeName: string): word.Shape
        getCount(): int
        addCallout(type: int, left: double, top: double, width: double, height: double, anchor: any): word.Shape
        addPicture(fileName: string, linkToFileObj: any, saveWithDocumentObj: any, leftObj: any, topObj: any, widthObj: any, heightObj: any, anchorObj: any): word.Shape
        addPolyline(safeArrayOfPoints: any, anchor: any): word.Shape
        addTextbox(orientation: int, left: double, top: double, width: double, height: double, anchor: any): word.Shape
        addTextEffect(presetTextEffect: int, text: string, fontName: string, fontSize: double, fontBold: int, fontItalic: int, left: double, top: double, anchor: any): word.Shape
        buildFreeform(editingType: int, x1: double, y1: double): word.FreeformBuilder
        selectAll(): void
        addCurve(safeArrayOfPoints: any, anchor: any): word.Shape
        addLabel(orientation: int, left: double, top: double, width: double, height: double, anchor: any): word.Shape
        addLine(beginX: double, beginY: double, endX: double, endY: double, anchor: any): word.Shape
        addShape(type: int, left: double, top: double, width: double, height: double, anchor: any): word.Shape
        addOLEControl(classType: any, left: any, top: any, width: any, height: any, anchor: any): word.Shape
        addOLEObject(classType: any, fileName: any, linkToFile: any, displayAsIcon: any, iconFileName: any, iconIndex: any, iconLabel: any, left: any, top: any, width: any, height: any, anchor: any): word.Shape
        getMShapes(): any
        checkDoc(): void
        getAllShapes(): word.Shape[]
        addCanvas(left: double, top: double, width: double, height: double, anchor: any): word.Shape
        addDiagram(type: int, left: float, top: float, width: float, height: float, anchor: any): word.Shape
        checkPicture(): void
    }
    export interface SmartTag {
        getName(): string
        getProperties(): word.CustomProperties
        delete(): void
        setName(pName: string): void
        getRange(): word.Range
        select(): void
        getXML(): string
        getDownloadURL(): string
        getXMLNode(): word.XMLNode
        getSmartTagActions(): word.SmartTagActions
    }
    export interface SmartTag {
        getName(): string
        getProperties(): word.CustomProperties
        delete(): void
        setName(pName: string): void
        getRange(): word.Range
        select(): void
        getXML(): string
        getDownloadURL(): string
        getXMLNode(): word.XMLNode
        getSmartTagActions(): word.SmartTagActions
    }
    export interface SmartTagAction {
        getName(): string
        execute(): void
        setName(pName: string): void
        getType(): int
        setExpandDocumentFragment(pExpandDocumentFragment: boolean): void
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
        isExpandDocumentFragment(): boolean
        getRadioGroupSelection(): int
        setRadioGroupSelection(pRadioGroupSelection: int): void
    }
    export interface SmartTagAction {
        getName(): string
        execute(): void
        setName(pName: string): void
        getType(): int
        setExpandDocumentFragment(pExpandDocumentFragment: boolean): void
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
        isExpandDocumentFragment(): boolean
        getRadioGroupSelection(): int
        setRadioGroupSelection(pRadioGroupSelection: int): void
    }
    export interface SmartTagActions {
        item(index: any): word.SmartTagAction
        getCount(): int
        reloadActions(): void
    }
    export interface SmartTagActions {
        item(index: any): word.SmartTagAction
        getCount(): int
        reloadActions(): void
    }
    export interface SmartTagRecognizer {
        getCaption(): string
        getFullName(): string
        setEnabled(pEnabled: boolean): void
        isEnabled(): boolean
        getProgID(): string
    }
    export interface SmartTagRecognizer {
        getCaption(): string
        getFullName(): string
        setEnabled(pEnabled: boolean): void
        isEnabled(): boolean
        getProgID(): string
    }
    export interface SmartTagRecognizers {
        item(index: any): word.SmartTagRecognizer
        getCount(): int
        reloadRecognizers(): void
    }
    export interface SmartTagRecognizers {
        item(index: any): word.SmartTagRecognizer
        getCount(): int
        reloadRecognizers(): void
    }
    export interface SmartTags {
        add(name: string, Range: any, Properties: any): word.SmartTag
        item(index: any): word.SmartTag
        getCount(): int
        smartTagsByType(name: string): word.SmartTags
    }
    export interface SmartTags {
        add(name: string, Range: any, Properties: any): word.SmartTag
        item(index: any): word.SmartTag
        getCount(): int
        smartTagsByType(name: string): word.SmartTags
    }
    export interface SmartTagType {
        getName(): string
        getSmartTagRecognizers(): word.SmartTagRecognizers
        getFriendlyName(): string
        getSmartTagActions(): word.SmartTagActions
    }
    export interface SmartTagType {
        getName(): string
        getSmartTagRecognizers(): word.SmartTagRecognizers
        getFriendlyName(): string
        getSmartTagActions(): word.SmartTagActions
    }
    export interface SmartTagTypes {
        item(index: any): word.SmartTagType
        getCount(): int
        reloadAll(): void
    }
    export interface SmartTagTypes {
        item(index: any): word.SmartTagType
        getCount(): int
        reloadAll(): void
    }
    export interface SortProperties {
        getSeparator(): char
        setSeparator(c: char): void
        getMTableSortProperties(): any
        setExcludeTitle(hasTitle: boolean): void
        setSortKey1(key1: int): void
        setSortKeyType1(sortType: int): void
        setAscending1(ascending: boolean): void
        setCaseSensitive(matchCase: boolean): void
        setSortKey2(key2: int): void
        setAscending2(ascending: boolean): void
        setSortKeyType2(sortType: int): void
        setSortKey3(key3: int): void
        setAscending3(ascending: boolean): void
        setSortKeyType3(sortType: int): void
        setSortSubKey3(subkey3: int): void
        isExcludeTitle(): boolean
        isCaseSensitive(): boolean
        setSortSubKey1(subkey1: int): void
        setSortSubKey2(subkey2: int): void
        setSortRowOnly(bool: boolean): void
        setSortOrientation(orientation: int): void
        getSortOrientation(): int
        setSortColumnOnly(bool: boolean): void
    }
    export interface SortProperties {
        getSeparator(): char
        setSeparator(c: char): void
        getMTableSortProperties(): any
        setExcludeTitle(hasTitle: boolean): void
        setSortKey1(key1: int): void
        setSortKeyType1(sortType: int): void
        setAscending1(ascending: boolean): void
        setCaseSensitive(matchCase: boolean): void
        setSortKey2(key2: int): void
        setAscending2(ascending: boolean): void
        setSortKeyType2(sortType: int): void
        setSortKey3(key3: int): void
        setAscending3(ascending: boolean): void
        setSortKeyType3(sortType: int): void
        setSortSubKey3(subkey3: int): void
        isExcludeTitle(): boolean
        isCaseSensitive(): boolean
        setSortSubKey1(subkey1: int): void
        setSortSubKey2(subkey2: int): void
        setSortRowOnly(bool: boolean): void
        setSortOrientation(orientation: int): void
        getSortOrientation(): int
        setSortColumnOnly(bool: boolean): void
    }
    export interface SpellingSuggestion {
        getName(): string
    }
    export interface SpellingSuggestion {
        getName(): string
    }
    export interface SpellingSuggestions {
        item(index: int): word.SpellingSuggestion
        getCount(): int
        getSpellingErrorType(): int
    }
    export interface SpellingSuggestions {
        item(index: int): word.SpellingSuggestion
        getCount(): int
        getSpellingErrorType(): int
    }
    export interface StoryRanges {
        item(index: int): word.Range
        getCount(): int
    }
    export interface StoryRanges {
        item(index: int): word.Range
        getCount(): int
    }
    export interface Style {
        delete(): void
        getType(): int
        getTable(): word.TableStyle
        isLocked(): boolean
        getFont(): word.Font
        setFont(pFont: word.Font): void
        getFrame(): word.Frame
        setBorders(pBorders: word.Borders): void
        getDescription(): string
        getBorders(): word.Borders
        getShading(): word.Shading
        setParagraphFormat(pParagraphFormat: word.ParagraphFormat): void
        getParagraphFormat(): word.ParagraphFormat
        modify(): void
        getNoProofing(): int
        setNoProofing(pNoProofing: int): void
        getLanguageIDFarEast(): int
        setLanguageIDFarEast(pLanguageIDFarEast: int): void
        getNameLocal(): string
        getListTemplate(): word.ListTemplate
        setLocked(pLocked: boolean): void
        getLanguageID(): int
        setLanguageID(pLanguageID: int): void
        getListLevelNumber(): int
        getMStyle(): any
        isInUse(): boolean
        isNoSpaceBetweenParagraphsOfSameStyle(): boolean
        setNoSpaceBetweenParagraphsOfSameStyle(pNoSpaceBetweenParagraphsOfSameStyle: boolean): void
        getBaseStyle(): any
        isBuiltIn(): boolean
        getLinkStyle(): any
        setBaseStyle(pBaseStyle: any): void
        setLinkStyle(pLinkStyle: any): void
        setNameLocal(pNameLocal: string): void
        isAutomaticallyUpdate(): boolean
        getNextParagraphStyle(): any
        setAutomaticallyUpdate(pAutomaticallyUpdate: boolean): void
        setNextParagraphStyle(pNextParagraphStyle: any): void
        LinkToListTemplate(listTemplate: word.ListTemplate, ListLevelNumber: any): void
    }
    export interface Style {
        delete(): void
        getType(): int
        getTable(): word.TableStyle
        isLocked(): boolean
        getFont(): word.Font
        setFont(pFont: word.Font): void
        getFrame(): word.Frame
        setBorders(pBorders: word.Borders): void
        getDescription(): string
        getBorders(): word.Borders
        getShading(): word.Shading
        setParagraphFormat(pParagraphFormat: word.ParagraphFormat): void
        getParagraphFormat(): word.ParagraphFormat
        modify(): void
        getNoProofing(): int
        setNoProofing(pNoProofing: int): void
        getLanguageIDFarEast(): int
        setLanguageIDFarEast(pLanguageIDFarEast: int): void
        getNameLocal(): string
        getListTemplate(): word.ListTemplate
        setLocked(pLocked: boolean): void
        getLanguageID(): int
        setLanguageID(pLanguageID: int): void
        getListLevelNumber(): int
        getMStyle(): any
        isInUse(): boolean
        isNoSpaceBetweenParagraphsOfSameStyle(): boolean
        setNoSpaceBetweenParagraphsOfSameStyle(pNoSpaceBetweenParagraphsOfSameStyle: boolean): void
        getBaseStyle(): any
        isBuiltIn(): boolean
        getLinkStyle(): any
        setBaseStyle(pBaseStyle: any): void
        setLinkStyle(pLinkStyle: any): void
        setNameLocal(pNameLocal: string): void
        isAutomaticallyUpdate(): boolean
        getNextParagraphStyle(): any
        setAutomaticallyUpdate(pAutomaticallyUpdate: boolean): void
        setNextParagraphStyle(pNextParagraphStyle: any): void
        LinkToListTemplate(listTemplate: word.ListTemplate, ListLevelNumber: any): void
    }
    export interface Styles {
        add(name: string, Type: int): word.Style
        item(index: any): word.Style
        getCount(): int
    }
    export interface Styles {
        add(name: string, Type: int): word.Style
        item(index: any): word.Style
        getCount(): int
    }
    export interface StyleSheet {
        getName(): string
        delete(): void
        setName(pName: string): void
        getType(): int
        getPath(): string
        getIndex(): int
        setType(pType: int): void
        setTitle(pTitle: string): void
        getTitle(): string
        move(precedence: int): void
        getFullName(): string
    }
    export interface StyleSheet {
        getName(): string
        delete(): void
        setName(pName: string): void
        getType(): int
        getPath(): string
        getIndex(): int
        setType(pType: int): void
        setTitle(pTitle: string): void
        getTitle(): string
        move(precedence: int): void
        getFullName(): string
    }
    export interface StyleSheets {
        add(fileName: string, linkType: int, title: string, precedence: int): word.StyleSheet
        item(index: any): word.StyleSheet
        getCount(): int
    }
    export interface StyleSheets {
        add(fileName: string, linkType: int, title: string, precedence: int): word.StyleSheet
        item(index: any): word.StyleSheet
        getCount(): int
    }
    export interface Subdocument {
        getName(): string
        split(range: word.Range): void
        delete(): void
        setName(name: string): void
        getPath(): string
        open(): word.Document
        isLocked(): boolean
        getRange(): word.Range
        setLocked(pLocked: boolean): void
        getLevel(): int
        isHasFile(): boolean
    }
    export interface Subdocument {
        getName(): string
        split(range: word.Range): void
        delete(): void
        setName(name: string): void
        getPath(): string
        open(): word.Document
        isLocked(): boolean
        getRange(): word.Range
        setLocked(pLocked: boolean): void
        getLevel(): int
        isHasFile(): boolean
    }
    export interface Subdocuments {
        delete(): void
        merge(firstSubdocument: any, lastSubdocument: any): void
        item(index: int): word.Subdocument
        getCount(): int
        select(): void
        isExpanded(): boolean
        setExpanded(pExpanded: boolean): void
        addFromFile(name: string, confirmConversions: any, readOnly: any, passwordDocument: any, passwordTemplate: any, revert: any, writePasswordDocument: any, writePasswordTemplate: any): word.Subdocument
        addFromRange(range: word.Range): word.Subdocument
    }
    export interface Subdocuments {
        delete(): void
        merge(firstSubdocument: any, lastSubdocument: any): void
        item(index: int): word.Subdocument
        getCount(): int
        select(): void
        isExpanded(): boolean
        setExpanded(pExpanded: boolean): void
        addFromFile(name: string, confirmConversions: any, readOnly: any, passwordDocument: any, passwordTemplate: any, revert: any, writePasswordDocument: any, writePasswordTemplate: any): word.Subdocument
        addFromRange(range: word.Range): word.Subdocument
    }
    export interface SynonymInfo {
        getWord(): string
        isFound(): boolean
        getAntonymList(): any
        getMeaningCount(): int
        getMeaningList(): any
        getSynonymList(meaning: any): any
        getPartOfSpeechList(): any
        getRelatedExpressionList(): any
        getRelatedWordList(): any
    }
    export interface SynonymInfo {
        getWord(): string
        isFound(): boolean
        getAntonymList(): any
        getMeaningCount(): int
        getMeaningList(): any
        getSynonymList(meaning: any): any
        getPartOfSpeechList(): any
        getRelatedExpressionList(): any
        getRelatedWordList(): any
    }
    export interface System {
        connect(path: string, drive: any, password: any): void
        setCursor(pCursor: int): void
        getCursor(): int
        getVersion(): string
        isMathCoprocessorInstalled(): boolean
        mSInfo(): void
        getComputerType(): string
        getCountryRegion(): int
        getFreeDiskSpace(): int
        getMacintoshName(): string
        getProcessorType(): string
        getProfileString(section: string, key: string): string
        setProfileString(section: string, key: string, pProfileString: string): void
        getHorizontalResolution(): int
        getLanguageDesignation(): string
        getOperatingSystem(): string
        getPrivateProfileString(fileName: string, section: string, key: string): string
        isQuickDrawInstalled(): boolean
        getVerticalResolution(): int
        setPrivateProfileString(fileName: string, section: string, key: string, pPrivateProfileString: string): void
    }
    export interface System {
        connect(path: string, drive: any, password: any): void
        setCursor(pCursor: int): void
        getCursor(): int
        getVersion(): string
        isMathCoprocessorInstalled(): boolean
        mSInfo(): void
        getComputerType(): string
        getCountryRegion(): int
        getFreeDiskSpace(): int
        getMacintoshName(): string
        getProcessorType(): string
        getProfileString(section: string, key: string): string
        setProfileString(section: string, key: string, pProfileString: string): void
        getHorizontalResolution(): int
        getLanguageDesignation(): string
        getOperatingSystem(): string
        getPrivateProfileString(fileName: string, section: string, key: string): string
        isQuickDrawInstalled(): boolean
        getVerticalResolution(): int
        setPrivateProfileString(fileName: string, section: string, key: string, pPrivateProfileString: string): void
    }
    export interface Table {
        split(beforeRow: any): word.Table
        delete(): void
        start(): void
        sort(excludeHeader: any, fieldNumber: any, sortFieldType: any, sortOrder: any, fieldNumber2: any, sortFieldType2: any, sortOrder2: any, fieldNumber3: any, sortFieldType3: any, sortOrder3: any, caseSensitive: any, bidiSort: any, ignoreThe: any, ignoreKashida: any, ignoreDiacritics: any, ignoreHe: any, languageID: any): void
        sort(excludeHeader: boolean, fieldNumber: any, sortFieldType: int, sortOrder: int, fieldNumber2: any, sortFieldType2: int, sortOrder2: int, fieldNumber3: any, sortFieldType3: int, sortOrder3: int, caseSensitive: boolean, bidiSort: boolean, ignoreThe: boolean, ignoreKashida: boolean, ignoreDiacritics: boolean, ignoreHe: boolean, languageID: int): void
        getID(): string
        end(): void
        setLeftPadding(pLeftPadding: double): void
        getLeftPadding(): double
        getNestingLevel(): int
        getColCount(): int
        getRowCount(): int
        getTables(): word.Tables
        access$0(arg0: word.Table, arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any): void
        getRange(): word.Range
        select(): void
        setPreferredWidth(pPreferredWidth: double): void
        getPreferredWidth(): double
        setPreferredWidthType(pPreferredWidthType: int): void
        getPreferredWidthType(): int
        setID(pID: string): void
        cell(row: int, column: int): word.Cell
        setBorders(pBorders: word.Borders): void
        getMTable(): any
        getBorders(): word.Borders
        setBottomPadding(pBottomPadding: double): void
        getBottomPadding(): double
        checkIsNull(): void
        getShading(): word.Shading
        setRightPadding(pRightPadding: double): void
        getRightPadding(): double
        setTopPadding(pTopPadding: double): void
        getTopPadding(): double
        setSpacing(pSpacing: double): void
        getSpacing(): double
        getColumns(): word.Columns
        getRows(): word.Rows
        getStyle(): any
        setStyle(pStyle: any): void
        autoFormat(format: any, applyBorders: any, applyShading: any, applyFont: any, applyColor: any, applyHeadingRows: any, applyLastRow: any, applyFirstColumn: any, applyLastColumn: any, autoFit: any): void
        sortAscending(): void
        sortDescending(): void
        convertToText(separator: any, nestedTables: any): word.Range
        convertToText(start: long, end: long, separator: any, nestedTables: any): word.Range
        checkRow(row: int, rows: int): int
        autoFitBehavior(behavior: int): void
        isAllowAutoFit(): boolean
        setAllowAutoFit(pAllowAutoFit: boolean): void
        isUniform(): boolean
        autoFormat2(format: any, applyBorders: any, applyShading: any, applyFont: any, applyColor: any, applyHeadingRows: any, applyLastRow: any, applyFirstColumn: any, applyLastColumn: any, autoFitObj: any): void
        checkColumn(col: int, columns: int): int
        updateAutoFormat(): void
        setTableName(tableName: string): void
        isAllowPageBreaks(): boolean
        setAllowPageBreaks(pAllowPageBreaks: boolean): void
        isApplyStyleFirstColumn(): boolean
        setApplyStyleFirstColumn(pApplyStyleFirstColumn: boolean): void
        isApplyStyleHeadingRows(): boolean
        setApplyStyleHeadingRows(pApplyStyleHeadingRows: boolean): void
        isApplyStyleLastColumn(): boolean
        setApplyStyleLastColumn(pApplyStyleLastColumn: boolean): void
        isApplyStyleLastRow(): boolean
        setApplyStyleLastRow(pApplyStyleLastRow: boolean): void
        getAutoFormatType(): int
        changeFormatType_To_MS(eioType: int): int
        changeFormatType_To_EIO(msType: int): int
        fieldNumberParseKey(fieldNumberObj: any): int
        getTableDirection(): int
        setTableDirection(pTableDirection: int): void
    }
    export interface Table {
        split(beforeRow: any): word.Table
        delete(): void
        start(): void
        sort(excludeHeader: any, fieldNumber: any, sortFieldType: any, sortOrder: any, fieldNumber2: any, sortFieldType2: any, sortOrder2: any, fieldNumber3: any, sortFieldType3: any, sortOrder3: any, caseSensitive: any, bidiSort: any, ignoreThe: any, ignoreKashida: any, ignoreDiacritics: any, ignoreHe: any, languageID: any): void
        sort(excludeHeader: boolean, fieldNumber: any, sortFieldType: int, sortOrder: int, fieldNumber2: any, sortFieldType2: int, sortOrder2: int, fieldNumber3: any, sortFieldType3: int, sortOrder3: int, caseSensitive: boolean, bidiSort: boolean, ignoreThe: boolean, ignoreKashida: boolean, ignoreDiacritics: boolean, ignoreHe: boolean, languageID: int): void
        getID(): string
        end(): void
        setLeftPadding(pLeftPadding: double): void
        getLeftPadding(): double
        getNestingLevel(): int
        getColCount(): int
        getRowCount(): int
        getTables(): word.Tables
        access$0(arg0: word.Table, arg1: any, arg2: any, arg3: any, arg4: any, arg5: any, arg6: any, arg7: any, arg8: any, arg9: any, arg10: any): void
        getRange(): word.Range
        select(): void
        setPreferredWidth(pPreferredWidth: double): void
        getPreferredWidth(): double
        setPreferredWidthType(pPreferredWidthType: int): void
        getPreferredWidthType(): int
        setID(pID: string): void
        cell(row: int, column: int): word.Cell
        setBorders(pBorders: word.Borders): void
        getMTable(): any
        getBorders(): word.Borders
        setBottomPadding(pBottomPadding: double): void
        getBottomPadding(): double
        checkIsNull(): void
        getShading(): word.Shading
        setRightPadding(pRightPadding: double): void
        getRightPadding(): double
        setTopPadding(pTopPadding: double): void
        getTopPadding(): double
        setSpacing(pSpacing: double): void
        getSpacing(): double
        getColumns(): word.Columns
        getRows(): word.Rows
        getStyle(): any
        setStyle(pStyle: any): void
        autoFormat(format: any, applyBorders: any, applyShading: any, applyFont: any, applyColor: any, applyHeadingRows: any, applyLastRow: any, applyFirstColumn: any, applyLastColumn: any, autoFit: any): void
        sortAscending(): void
        sortDescending(): void
        convertToText(separator: any, nestedTables: any): word.Range
        convertToText(start: long, end: long, separator: any, nestedTables: any): word.Range
        checkRow(row: int, rows: int): int
        autoFitBehavior(behavior: int): void
        isAllowAutoFit(): boolean
        setAllowAutoFit(pAllowAutoFit: boolean): void
        isUniform(): boolean
        autoFormat2(format: any, applyBorders: any, applyShading: any, applyFont: any, applyColor: any, applyHeadingRows: any, applyLastRow: any, applyFirstColumn: any, applyLastColumn: any, autoFitObj: any): void
        checkColumn(col: int, columns: int): int
        updateAutoFormat(): void
        setTableName(tableName: string): void
        isAllowPageBreaks(): boolean
        setAllowPageBreaks(pAllowPageBreaks: boolean): void
        isApplyStyleFirstColumn(): boolean
        setApplyStyleFirstColumn(pApplyStyleFirstColumn: boolean): void
        isApplyStyleHeadingRows(): boolean
        setApplyStyleHeadingRows(pApplyStyleHeadingRows: boolean): void
        isApplyStyleLastColumn(): boolean
        setApplyStyleLastColumn(pApplyStyleLastColumn: boolean): void
        isApplyStyleLastRow(): boolean
        setApplyStyleLastRow(pApplyStyleLastRow: boolean): void
        getAutoFormatType(): int
        changeFormatType_To_MS(eioType: int): int
        changeFormatType_To_EIO(msType: int): int
        fieldNumberParseKey(fieldNumberObj: any): int
        getTableDirection(): int
        setTableDirection(pTableDirection: int): void
    }
    export interface TableOfAuthorities {
        update(): void
        delete(): void
        getSeparator(): string
        getRange(): word.Range
        getBookmark(): string
        setSeparator(pSeparator: string): void
        getCategory(): int
        getTabLeader(): int
        setTabLeader(pTabLeader: int): void
        isPassim(): boolean
        setBookmark(pBookmark: string): void
        setCategory(pCategory: int): void
        setPassim(pPassim: boolean): void
        getEntrySeparator(): string
        isIncludeCategoryHeader(): boolean
        getIncludeSequenceName(): string
        isKeepEntryFormatting(): boolean
        getPageNumberSeparator(): string
        getPageRangeSeparator(): string
        setEntrySeparator(pEntrySeparator: string): void
        setIncludeCategoryHeader(pIncludeCategoryHeader: boolean): void
        setIncludeSequenceName(pIncludeSequenceName: string): void
        setKeepEntryFormatting(pKeepEntryFormatting: boolean): void
        setPageNumberSeparator(pPageNumberSeparator: string): void
        setPageRangeSeparator(pPageRangeSeparator: string): void
    }
    export interface TableOfAuthorities {
        update(): void
        delete(): void
        getSeparator(): string
        getRange(): word.Range
        getBookmark(): string
        setSeparator(pSeparator: string): void
        getCategory(): int
        getTabLeader(): int
        setTabLeader(pTabLeader: int): void
        isPassim(): boolean
        setBookmark(pBookmark: string): void
        setCategory(pCategory: int): void
        setPassim(pPassim: boolean): void
        getEntrySeparator(): string
        isIncludeCategoryHeader(): boolean
        getIncludeSequenceName(): string
        isKeepEntryFormatting(): boolean
        getPageNumberSeparator(): string
        getPageRangeSeparator(): string
        setEntrySeparator(pEntrySeparator: string): void
        setIncludeCategoryHeader(pIncludeCategoryHeader: boolean): void
        setIncludeSequenceName(pIncludeSequenceName: string): void
        setKeepEntryFormatting(pKeepEntryFormatting: boolean): void
        setPageNumberSeparator(pPageNumberSeparator: string): void
        setPageRangeSeparator(pPageRangeSeparator: string): void
    }
    export interface TableOfAuthoritiesCategory {
        getName(): string
        setName(pName: string): void
        getIndex(): int
    }
    export interface TableOfAuthoritiesCategory {
        getName(): string
        setName(pName: string): void
        getIndex(): int
    }
    export interface TableOfContents {
        update(): void
        delete(): void
        getRange(): word.Range
        getTabLeader(): int
        setTabLeader(pTabLeader: int): void
        setRightAlignPageNumbers(pRightAlignPageNumbers: boolean): void
        isRightAlignPageNumbers(): boolean
        getHeadingStyles(): word.HeadingStyles
        getTableID(): string
        setTableID(pTableID: string): void
        isUseFields(): boolean
        setUseFields(pUseFields: boolean): void
        isUseHyperlinks(): boolean
        setUseHyperlinks(pUseHyperlinks: boolean): void
        isHidePageNumbersInWeb(): boolean
        setHidePageNumbersInWeb(pHidePageNumbersInWeb: boolean): void
        isIncludePageNumbers(): boolean
        setIncludePageNumbers(pIncludePageNumbers: boolean): void
        getLowerHeadingLevel(): int
        setLowerHeadingLevel(pLowerHeadingLevel: int): void
        getUpperHeadingLevel(): int
        setUpperHeadingLevel(pUpperHeadingLevel: int): void
        isUseHeadingStyles(): boolean
        setUseHeadingStyles(pUseHeadingStyles: boolean): void
        updatePageNumbers(): void
    }
    export interface TableOfContents {
        update(): void
        delete(): void
        getRange(): word.Range
        getTabLeader(): int
        setTabLeader(pTabLeader: int): void
        setRightAlignPageNumbers(pRightAlignPageNumbers: boolean): void
        isRightAlignPageNumbers(): boolean
        getHeadingStyles(): word.HeadingStyles
        getTableID(): string
        setTableID(pTableID: string): void
        isUseFields(): boolean
        setUseFields(pUseFields: boolean): void
        isUseHyperlinks(): boolean
        setUseHyperlinks(pUseHyperlinks: boolean): void
        isHidePageNumbersInWeb(): boolean
        setHidePageNumbersInWeb(pHidePageNumbersInWeb: boolean): void
        isIncludePageNumbers(): boolean
        setIncludePageNumbers(pIncludePageNumbers: boolean): void
        getLowerHeadingLevel(): int
        setLowerHeadingLevel(pLowerHeadingLevel: int): void
        getUpperHeadingLevel(): int
        setUpperHeadingLevel(pUpperHeadingLevel: int): void
        isUseHeadingStyles(): boolean
        setUseHeadingStyles(pUseHeadingStyles: boolean): void
        updatePageNumbers(): void
    }
    export interface TableOfFigures {
        update(): void
        delete(): void
        setCaption(pCaption: string): void
        getCaption(): string
        getRange(): word.Range
        getTabLeader(): int
        setTabLeader(pTabLeader: int): void
        setRightAlignPageNumbers(pRightAlignPageNumbers: boolean): void
        isRightAlignPageNumbers(): boolean
        getHeadingStyles(): word.HeadingStyles
        getTableID(): string
        setTableID(pTableID: string): void
        isUseFields(): boolean
        setUseFields(pUseFields: boolean): void
        isUseHyperlinks(): boolean
        setUseHyperlinks(pUseHyperlinks: boolean): void
        isIncludeLabel(): boolean
        setIncludeLabel(pIncludeLabel: boolean): void
        isHidePageNumbersInWeb(): boolean
        setHidePageNumbersInWeb(pHidePageNumbersInWeb: boolean): void
        isIncludePageNumbers(): boolean
        setIncludePageNumbers(pIncludePageNumbers: boolean): void
        getLowerHeadingLevel(): int
        setLowerHeadingLevel(pLowerHeadingLevel: int): void
        getUpperHeadingLevel(): int
        setUpperHeadingLevel(pUpperHeadingLevel: int): void
        isUseHeadingStyles(): boolean
        setUseHeadingStyles(pUseHeadingStyles: boolean): void
        updatePageNumbers(): void
    }
    export interface TableOfFigures {
        update(): void
        delete(): void
        setCaption(pCaption: string): void
        getCaption(): string
        getRange(): word.Range
        getTabLeader(): int
        setTabLeader(pTabLeader: int): void
        setRightAlignPageNumbers(pRightAlignPageNumbers: boolean): void
        isRightAlignPageNumbers(): boolean
        getHeadingStyles(): word.HeadingStyles
        getTableID(): string
        setTableID(pTableID: string): void
        isUseFields(): boolean
        setUseFields(pUseFields: boolean): void
        isUseHyperlinks(): boolean
        setUseHyperlinks(pUseHyperlinks: boolean): void
        isIncludeLabel(): boolean
        setIncludeLabel(pIncludeLabel: boolean): void
        isHidePageNumbersInWeb(): boolean
        setHidePageNumbersInWeb(pHidePageNumbersInWeb: boolean): void
        isIncludePageNumbers(): boolean
        setIncludePageNumbers(pIncludePageNumbers: boolean): void
        getLowerHeadingLevel(): int
        setLowerHeadingLevel(pLowerHeadingLevel: int): void
        getUpperHeadingLevel(): int
        setUpperHeadingLevel(pUpperHeadingLevel: int): void
        isUseHeadingStyles(): boolean
        setUseHeadingStyles(pUseHeadingStyles: boolean): void
        updatePageNumbers(): void
    }
    export interface Tables {
        add(range: word.Range, numRows: int, numColumns: int, defaultTableBehavior: any, autoFitBehavior: any): word.Table
        item(index: int): word.Table
        item(tableName: string): word.Table
        getCount(): int
        getNestingLevel(): int
        getActiveTable(): word.Table
        getMTables(): any
        getTablesParent(): any
    }
    export interface Tables {
        add(range: word.Range, numRows: int, numColumns: int, defaultTableBehavior: any, autoFitBehavior: any): word.Table
        item(index: int): word.Table
        item(tableName: string): word.Table
        getCount(): int
        getNestingLevel(): int
        getActiveTable(): word.Table
        getMTables(): any
        getTablesParent(): any
    }
    export interface TablesOfAuthorities {
        add(range: word.Range, category: any, bookmark: any, passim: any, keepEntryFormatting: any, separator: any, includeSequenceName: any, entrySeparator: any, pageRangeSeparator: any, includeCategoryHeader: any, pageNumberSeparator: any): word.TableOfAuthorities
        item(index: int): word.TableOfAuthorities
        getCount(): int
        getFormat(): int
        setFormat(pFormat: int): void
        markAllCitations(shortCitation: string, intCitation: any, intCitationAutoText: any, category: any): void
        markCitation(range: word.Range, shortCitation: string, intCitation: any, intCitationAutoText: any, category: any): word.Field
        nextCitation(shortCitation: string): void
    }
    export interface TablesOfAuthorities {
        add(range: word.Range, category: any, bookmark: any, passim: any, keepEntryFormatting: any, separator: any, includeSequenceName: any, entrySeparator: any, pageRangeSeparator: any, includeCategoryHeader: any, pageNumberSeparator: any): word.TableOfAuthorities
        item(index: int): word.TableOfAuthorities
        getCount(): int
        getFormat(): int
        setFormat(pFormat: int): void
        markAllCitations(shortCitation: string, intCitation: any, intCitationAutoText: any, category: any): void
        markCitation(range: word.Range, shortCitation: string, intCitation: any, intCitationAutoText: any, category: any): word.Field
        nextCitation(shortCitation: string): void
    }
    export interface TablesOfAuthoritiesCategories {
        item(index: any): word.TableOfAuthoritiesCategory
        getCount(): int
    }
    export interface TablesOfAuthoritiesCategories {
        item(index: any): word.TableOfAuthoritiesCategory
        getCount(): int
    }
    export interface TablesOfContents {
        add(range: word.Range, useHeadingStyles: any, upperHeadingLevel: any, lowerHeadingLevel: any, useFields: any, tableID: any, rightAlignPageNumbers: any, includePageNumbers: any, addedStyles: any, useHyperlinks: any, hidePageNumbersInWeb: any, useOutlineLevels: any): word.TableOfContents
        item(index: int): word.TableOfContents
        getCount(): int
        getFormat(): int
        setFormat(pFormat: int): void
        markEntry(range: word.Range, entry: any, entryAutoText: any, tableID: any, level: any): word.Field
    }
    export interface TablesOfContents {
        add(range: word.Range, useHeadingStyles: any, upperHeadingLevel: any, lowerHeadingLevel: any, useFields: any, tableID: any, rightAlignPageNumbers: any, includePageNumbers: any, addedStyles: any, useHyperlinks: any, hidePageNumbersInWeb: any, useOutlineLevels: any): word.TableOfContents
        item(index: int): word.TableOfContents
        getCount(): int
        getFormat(): int
        setFormat(pFormat: int): void
        markEntry(range: word.Range, entry: any, entryAutoText: any, tableID: any, level: any): word.Field
    }
    export interface TablesOfFigures {
        add(range: word.Range, caption: any, includeLabel: any, useHeadingStyles: any, upperHeadingLevel: any, lowerHeadingLevel: any, useFields: any, tableID: any, rightAlignPageNumbers: any, includePageNumbers: any, addedStyles: any, useHyperlinks: any, hidePageNumbersInWeb: any): word.TableOfFigures
        item(index: int): word.TableOfFigures
        getCount(): int
        getFormat(): int
        setFormat(pFormat: int): void
        markEntry(range: word.Range, entry: any, entryAutoText: any, tableID: any, level: any): word.Field
    }
    export interface TablesOfFigures {
        add(range: word.Range, caption: any, includeLabel: any, useHeadingStyles: any, upperHeadingLevel: any, lowerHeadingLevel: any, useFields: any, tableID: any, rightAlignPageNumbers: any, includePageNumbers: any, addedStyles: any, useHyperlinks: any, hidePageNumbersInWeb: any): word.TableOfFigures
        item(index: int): word.TableOfFigures
        getCount(): int
        getFormat(): int
        setFormat(pFormat: int): void
        markEntry(range: word.Range, entry: any, entryAutoText: any, tableID: any, level: any): word.Field
    }
    export interface TableStyle {
        setLeftPadding(pLeftPadding: double): void
        getLeftPadding(): double
        setBorders(pBorders: word.Borders): void
        getBorders(): word.Borders
        setBottomPadding(pBottomPadding: double): void
        getBottomPadding(): double
        getShading(): word.Shading
        setRightPadding(pRightPadding: double): void
        getRightPadding(): double
        setTopPadding(pTopPadding: double): void
        getTopPadding(): double
        setSpacing(pSpacing: double): void
        getSpacing(): double
        getAlignment(): int
        setAlignment(pAlignment: int): void
        getLeftIndent(): double
        setLeftIndent(pLeftIndent: double): void
        getColumnStripe(): int
        getRowStripe(): int
        setColumnStripe(pColumnStripe: int): void
        setRowStripe(pRowStripe: int): void
        Condition(conditionCode: int): word.ConditionalStyle
        isAllowPageBreaks(): boolean
        setAllowPageBreaks(pAllowPageBreaks: boolean): void
        getAllowBreakAcrossPage(): int
        setAllowBreakAcrossPage(pAllowBreakAcrossPage: int): void
        getTableDirection(): int
        setTableDirection(pTableDirection: int): void
    }
    export interface TableStyle {
        setLeftPadding(pLeftPadding: double): void
        getLeftPadding(): double
        setBorders(pBorders: word.Borders): void
        getBorders(): word.Borders
        setBottomPadding(pBottomPadding: double): void
        getBottomPadding(): double
        getShading(): word.Shading
        setRightPadding(pRightPadding: double): void
        getRightPadding(): double
        setTopPadding(pTopPadding: double): void
        getTopPadding(): double
        setSpacing(pSpacing: double): void
        getSpacing(): double
        getAlignment(): int
        setAlignment(pAlignment: int): void
        getLeftIndent(): double
        setLeftIndent(pLeftIndent: double): void
        getColumnStripe(): int
        getRowStripe(): int
        setColumnStripe(pColumnStripe: int): void
        setRowStripe(pRowStripe: int): void
        Condition(conditionCode: int): word.ConditionalStyle
        isAllowPageBreaks(): boolean
        setAllowPageBreaks(pAllowPageBreaks: boolean): void
        getAllowBreakAcrossPage(): int
        setAllowBreakAcrossPage(pAllowBreakAcrossPage: int): void
        getTableDirection(): int
        setTableDirection(pTableDirection: int): void
    }
    export interface TabStop {
        clear(): void
        apply(): void
        getPrevious(): word.TabStop
        getNext(): word.TabStop
        setPosition(pPosition: double): void
        getPosition(): double
        getAlignment(): int
        setAlignment(pAlignment: int): void
        isCustomTab(): boolean
        getLeader(): int
        setLeader(pLeader: int): void
    }
    export interface TabStop {
        clear(): void
        apply(): void
        getPrevious(): word.TabStop
        getNext(): word.TabStop
        setPosition(pPosition: double): void
        getPosition(): double
        getAlignment(): int
        setAlignment(pAlignment: int): void
        isCustomTab(): boolean
        getLeader(): int
        setLeader(pLeader: int): void
    }
    export interface TabStops {
        add(position: double, alignment: any, leader: any): word.TabStop
        before(position: double): word.TabStop
        after(position: double): word.TabStop
        item(index: any): word.TabStop
        getCount(): int
        getMtabs(): any
        clearAll(): void
    }
    export interface TabStops {
        add(position: double, alignment: any, leader: any): word.TabStop
        before(position: double): word.TabStop
        after(position: double): word.TabStop
        item(index: any): word.TabStop
        getCount(): int
        getMtabs(): any
        clearAll(): void
    }
    export interface Task {
        getName(): string
        close(): void
        resize(width: int, height: int): void
        getWidth(): int
        getLeft(): int
        setLeft(pLeft: int): void
        setTop(pTop: int): void
        getTop(): int
        setWidth(pWidth: int): void
        activate(wait: any): void
        getHeight(): int
        setHeight(pHeight: int): void
        isVisible(): boolean
        setVisible(pVisible: boolean): void
        getWindowState(): int
        setWindowState(pWindowState: int): void
        move(left: int, top: int): void
        sendWindowMessage(message: int, wParam: int, lParam: int): void
    }
    export interface Task {
        getName(): string
        close(): void
        resize(width: int, height: int): void
        getWidth(): int
        getLeft(): int
        setLeft(pLeft: int): void
        setTop(pTop: int): void
        getTop(): int
        setWidth(pWidth: int): void
        activate(wait: any): void
        getHeight(): int
        setHeight(pHeight: int): void
        isVisible(): boolean
        setVisible(pVisible: boolean): void
        getWindowState(): int
        setWindowState(pWindowState: int): void
        move(left: int, top: int): void
        sendWindowMessage(message: int, wParam: int, lParam: int): void
    }
    export interface TaskPane {
        isVisible(): boolean
        setVisible(pVisible: boolean): void
    }
    export interface TaskPane {
        isVisible(): boolean
        setVisible(pVisible: boolean): void
    }
    export interface TaskPanes {
        item(index: int): word.TaskPane
        getCount(): int
    }
    export interface TaskPanes {
        item(index: int): word.TaskPane
        getCount(): int
    }
    export interface Tasks {
        exists(name: string): boolean
        item(index: any): word.Task
        getCount(): int
        exitWindows(): void
    }
    export interface Tasks {
        exists(name: string): boolean
        item(index: any): word.Task
        getCount(): int
        exitWindows(): void
    }
    export interface Template {
        getName(): string
        save(): void
        getType(): int
        getPath(): string
        getFullName(): string
        getFarEastLineBreakLevel(): int
        setFarEastLineBreakLevel(farLevel: int): void
        getJustificationMode(): int
        setJustificationMode(juMode: int): void
        isKerningByAlgorithm(): boolean
        setKerningByAlgorithm(kernBy: boolean): void
        getNoLineBreakAfter(): string
        setNoLineBreakAfter(nlineAfter: string): void
        getNoLineBreakBefore(): string
        setNoLineBreakBefore(nLineBefore: string): void
        getNoProofing(): int
        setNoProofing(nProof: int): void
        getLanguageIDFarEast(): int
        setLanguageIDFarEast(langEast: int): void
        getListTemplates(): word.ListTemplates
        isSaved(): boolean
        setSaved(save: boolean): void
        getAutoTextEntries(): word.AutoTextEntries
        openAsDocument(): word.Document
        getBuiltInDocumentProperties(): any
        getLanguageID(): int
        setLanguageID(langID: int): void
        getCustomDocumentProperties(): office.DocumentProperties
        getFarEastLineBreakLanguage(): int
        setFarEastLineBreakLanguage(farLang: int): void
    }
    export interface Template {
        getName(): string
        save(): void
        getType(): int
        getPath(): string
        getFullName(): string
        getFarEastLineBreakLevel(): int
        setFarEastLineBreakLevel(farLevel: int): void
        getJustificationMode(): int
        setJustificationMode(juMode: int): void
        isKerningByAlgorithm(): boolean
        setKerningByAlgorithm(kernBy: boolean): void
        getNoLineBreakAfter(): string
        setNoLineBreakAfter(nlineAfter: string): void
        getNoLineBreakBefore(): string
        setNoLineBreakBefore(nLineBefore: string): void
        getNoProofing(): int
        setNoProofing(nProof: int): void
        getLanguageIDFarEast(): int
        setLanguageIDFarEast(langEast: int): void
        getListTemplates(): word.ListTemplates
        isSaved(): boolean
        setSaved(save: boolean): void
        getAutoTextEntries(): word.AutoTextEntries
        openAsDocument(): word.Document
        getBuiltInDocumentProperties(): any
        getLanguageID(): int
        setLanguageID(langID: int): void
        getCustomDocumentProperties(): office.DocumentProperties
        getFarEastLineBreakLanguage(): int
        setFarEastLineBreakLanguage(farLang: int): void
    }
    export interface Templates {
        init(): void
        getCount(): int
        getItem(index: any): word.Template
        getNormalTemplate(): word.Template
    }
    export interface Templates {
        init(): void
        getCount(): int
        getItem(index: any): word.Template
        getNormalTemplate(): word.Template
    }
    export interface TextColumn {
        getWidth(): double
        setWidth(w: double): void
        getSpaceAfter(): double
        setSpaceAfter(sa: double): void
    }
    export interface TextColumn {
        getWidth(): double
        setWidth(w: double): void
        getSpaceAfter(): double
        setSpaceAfter(sa: double): void
    }
    export interface TextColumns {
        add(width: any, spacing: any, evenlyspaced: any): word.TextColumn
        getCount(): int
        getItem(index: int): word.TextColumn
        getWidth(): double
        setWidth(w: double): void
        setSpacing(sp: int): void
        getSpacing(): double
        setCount(numColumns: int): void
        setEvenlySpace(es: int): void
        getEvenlySpace(): int
        setFlowDirection(fd: int): void
        getFlowDirection(): int
        setLineBetween(lb: int): void
        getLineBetween(): int
    }
    export interface TextColumns {
        add(width: any, spacing: any, evenlyspaced: any): word.TextColumn
        getCount(): int
        getItem(index: int): word.TextColumn
        getWidth(): double
        setWidth(w: double): void
        setSpacing(sp: int): void
        getSpacing(): double
        setCount(numColumns: int): void
        setEvenlySpace(es: int): void
        getEvenlySpace(): int
        setFlowDirection(fd: int): void
        getFlowDirection(): int
        setLineBetween(lb: int): void
        getLineBetween(): int
    }
    export interface TextEffectFormat {
        getText(): string
        setFontSize(fs: double): void
        setText(tx: string): void
        getAlignment(): int
        setAlignment(align: int): void
        getFontName(): string
        setFontName(name: string): void
        setNormalizedHeight(nh: int): void
        getNormalizedHeight(): int
        setPresetTextEffect(pte: int): void
        getPresetTextEffect(): int
        toggleVerticalText(): void
        setFontBold(v: int): void
        getFontBold(): int
        setFontItalic(v: int): void
        getFontItalic(): int
        getFontSize(): double
        setKernedPairs(kp: int): void
        isKernedPairs(): int
        setPresetShape(ps: int): void
        getPresetShape(): int
        setRotatedChars(rc: int): void
        getRotatedChars(): int
        setTracking(tk: double): void
        getTracking(): double
    }
    export interface TextEffectFormat {
        getText(): string
        setFontSize(fs: double): void
        setText(tx: string): void
        getAlignment(): int
        setAlignment(align: int): void
        getFontName(): string
        setFontName(name: string): void
        setNormalizedHeight(nh: int): void
        getNormalizedHeight(): int
        setPresetTextEffect(pte: int): void
        getPresetTextEffect(): int
        toggleVerticalText(): void
        setFontBold(v: int): void
        getFontBold(): int
        setFontItalic(v: int): void
        getFontItalic(): int
        getFontSize(): double
        setKernedPairs(kp: int): void
        isKernedPairs(): int
        setPresetShape(ps: int): void
        getPresetShape(): int
        setRotatedChars(rc: int): void
        getRotatedChars(): int
        setTracking(tk: double): void
        getTracking(): double
    }
    export interface TextFrame {
        getName(): string
        getOrientation(): int
        getPrevious(): word.TextFrame
        getNext(): word.TextFrame
        setAutoSize(as: int): void
        setWordWrap(wr: int): void
        setOrientation(ori: int): void
        getTextRange(): word.Range
        getWordWrap(): int
        hasText(): int
        setNext(textFrame: word.TextFrame): void
        getMarginLeft(): double
        setMarginLeft(ml: double): void
        getMarginRight(): double
        setMarginRight(mr: double): void
        getMarginBottom(): double
        setMarginBottom(mb: double): void
        getMarginTop(): double
        setMarginTop(mt: double): void
        breakForwardLink(): void
        validLinkTarget(target: word.TextFrame): boolean
        getContainingRange(): word.Range
        isValidLinkTarget(tf: word.TextFrame): boolean
        setHorizontalAnchor(v: int): void
        getHorizontalAnchor(): int
        setVerticalAnchor(v: int): void
        getVerticalAnchor(): int
        getAutoSize(): int
        isOverflowing(): boolean
        setWarpFormat(v: int): void
        getWarpFormat(): int
        deleteText(): void
    }
    export interface TextFrame {
        getName(): string
        getOrientation(): int
        getPrevious(): word.TextFrame
        getNext(): word.TextFrame
        setAutoSize(as: int): void
        setWordWrap(wr: int): void
        setOrientation(ori: int): void
        getTextRange(): word.Range
        getWordWrap(): int
        hasText(): int
        setNext(textFrame: word.TextFrame): void
        getMarginLeft(): double
        setMarginLeft(ml: double): void
        getMarginRight(): double
        setMarginRight(mr: double): void
        getMarginBottom(): double
        setMarginBottom(mb: double): void
        getMarginTop(): double
        setMarginTop(mt: double): void
        breakForwardLink(): void
        validLinkTarget(target: word.TextFrame): boolean
        getContainingRange(): word.Range
        isValidLinkTarget(tf: word.TextFrame): boolean
        setHorizontalAnchor(v: int): void
        getHorizontalAnchor(): int
        setVerticalAnchor(v: int): void
        getVerticalAnchor(): int
        getAutoSize(): int
        isOverflowing(): boolean
        setWarpFormat(v: int): void
        getWarpFormat(): int
        deleteText(): void
    }
    export interface TextInput {
        clear(): void
        getDefault(): string
        setDefault(s: string): void
        isValid(): boolean
        getWidth(): int
        setWidth(w: int): void
        EditType(type: int, defText: string, fm: string, isEnabled: any): void
        getFormat(): string
        getTextType(): int
    }
    export interface TextInput {
        clear(): void
        getDefault(): string
        setDefault(s: string): void
        isValid(): boolean
        getWidth(): int
        setWidth(w: int): void
        EditType(type: int, defText: string, fm: string, isEnabled: any): void
        getFormat(): string
        getTextType(): int
    }
    export interface TextRetrievalMode {
        getViewType(): int
        getDuplicate(): word.TextRetrievalMode
        setViewType(vt: int): void
        isIncludeFieldCodes(): boolean
        isIncludeHiddenText(): boolean
        setIncludeFieldCodes(ifc: boolean): void
        setIncludeHiddenText(iht: boolean): void
    }
    export interface TextRetrievalMode {
        getViewType(): int
        getDuplicate(): word.TextRetrievalMode
        setViewType(vt: int): void
        isIncludeFieldCodes(): boolean
        isIncludeHiddenText(): boolean
        setIncludeFieldCodes(ifc: boolean): void
        setIncludeHiddenText(iht: boolean): void
    }
    export interface ThreeDFormat {
        setVisible(vis: int): void
        getVisible(): int
        setDepth(pDepth: double): void
        getDepth(): double
        getExtrusionColor(): word.ColorFormat
        setExtrusionColorType(ect: int): void
        getExtrusionColorType(): int
        setPresetMaterial(pm: int): void
        getPresetMaterial(): int
        incrementRotationX(irX: float): void
        incrementRotationY(irY: float): void
        SetExtrusionDirection(presetExtrusionDirection: int): void
        getPresetThreeDFormat(): int
        setPresetExtrusionDirection(ped: int): void
        getPresetExtrusionDirection(): int
        setPresetLightingDirection(pld: int): void
        getPresetLightingDirection(): int
        setPresetLightingSoftness(pls: int): void
        getPresetLightingSoftness(): int
        setPerspective(pers: int): void
        getPerspective(): int
        setRotationX(rotX: double): void
        getRotationX(): double
        setRotationY(rotY: double): void
        getRotationY(): double
        resetRotation(): void
        setThreeDFormat(presetThreeDFormat: int): void
    }
    export interface ThreeDFormat {
        setVisible(vis: int): void
        getVisible(): int
        setDepth(pDepth: double): void
        getDepth(): double
        getExtrusionColor(): word.ColorFormat
        setExtrusionColorType(ect: int): void
        getExtrusionColorType(): int
        setPresetMaterial(pm: int): void
        getPresetMaterial(): int
        incrementRotationX(irX: float): void
        incrementRotationY(irY: float): void
        SetExtrusionDirection(presetExtrusionDirection: int): void
        getPresetThreeDFormat(): int
        setPresetExtrusionDirection(ped: int): void
        getPresetExtrusionDirection(): int
        setPresetLightingDirection(pld: int): void
        getPresetLightingDirection(): int
        setPresetLightingSoftness(pls: int): void
        getPresetLightingSoftness(): int
        setPerspective(pers: int): void
        getPerspective(): int
        setRotationX(rotX: double): void
        getRotationX(): double
        setRotationY(rotY: double): void
        getRotationY(): double
        resetRotation(): void
        setThreeDFormat(presetThreeDFormat: int): void
    }
    export interface TwoInitialCapsException {
        getName(): string
        delete(): void
        getIndex(): int
    }
    export interface TwoInitialCapsException {
        getName(): string
        delete(): void
        getIndex(): int
    }
    export interface TwoInitialCapsExceptions {
        add(name: string): word.TwoInitialCapsException
        item(index: any): word.TwoInitialCapsException
        getCount(): int
        getAllTwoInitialCapsExceptions(): any
        getTwoInitialCapsExceptionName(): string
        setTwoInitialCapsExceptionName(twoInitialCapsException: string): void
    }
    export interface TwoInitialCapsExceptions {
        add(name: string): word.TwoInitialCapsException
        item(index: any): word.TwoInitialCapsException
        getCount(): int
        getAllTwoInitialCapsExceptions(): any
        getTwoInitialCapsExceptionName(): string
        setTwoInitialCapsExceptionName(twoInitialCapsException: string): void
    }
    export interface Variable {
        getName(check: boolean): string
        getName(): string
        getValue(check: boolean): string
        getValue(): string
        delete(): void
        setValue(va: string): void
        getIndex(): int
        checkValid(): void
    }
    export interface Variable {
        getName(check: boolean): string
        getName(): string
        getValue(check: boolean): string
        getValue(): string
        delete(): void
        setValue(va: string): void
        getIndex(): int
        checkValid(): void
    }
    export interface Variables {
        add(name: string, obj: any): word.Variable
        init(parent: word.Document): void
        delete(var: word.Variable): void
        item(obj: any): word.Variable
        getIndex(var: word.Variable): int
        getCount(): int
        mustSave(): void
    }
    export interface Variables {
        add(name: string, obj: any): word.Variable
        init(parent: word.Document): void
        delete(var: word.Variable): void
        item(obj: any): word.Variable
        getIndex(var: word.Variable): int
        getCount(): int
        mustSave(): void
    }
    export interface Version {
        delete(): void
        open(): word.Document
        getDate(): any
        getIndex(): int
        getComment(): string
        initInfo(versionInfo: string): void
        getVersionDate(): double
        getSavedBy(): string
    }
    export interface Version {
        delete(): void
        open(): word.Document
        getDate(): any
        getIndex(): int
        getComment(): string
        initInfo(versionInfo: string): void
        getVersionDate(): double
        getSavedBy(): string
    }
    export interface Versions {
        save(obj: any): void
        item(index: int): word.Version
        getCount(): int
        getMWorkbook(): any
        setAutoVersion(av: int): void
        getAutoVersion(): int
    }
    export interface Versions {
        save(obj: any): void
        item(index: int): word.Version
        getCount(): int
        getMWorkbook(): any
        setAutoVersion(av: int): void
        getAutoVersion(): int
    }
    export interface View {
        getType(): int
        setType(t: int): void
        setFullScreen(fs: boolean): void
        isFullScreen(): boolean
        getZoom(): word.Zoom
        setShowRevisionsAndComments(srac: boolean): void
        isShowRevisionsAndComments(): boolean
        setTableGridlines(tgl: boolean): void
        setShadeEditableRanges(ser: int): void
        getShadeEditableRanges(): int
        getRevisionsView(): int
        setRevisionsView(rv: int): void
        getReviewers(): word.Reviewers
        isReadingLayout(): boolean
        setReadingLayout(rl: boolean): void
        setSeekView(sv: int): void
        getSeekView(): int
        isTableGridlines(): boolean
        isDraft(): boolean
        isShowAll(): boolean
        setShowAll(sa: boolean): void
        isDisplayBackgrounds(): boolean
        isDisplayPageBoundaries(): boolean
        isDisplaySmartTags(): boolean
        isMailMergeDataView(): boolean
        isShowFormatChanges(): boolean
        isShowInkAnnotations(): boolean
        isShowMainTextLayer(): boolean
        isShowObjectAnchors(): boolean
        isShowOptionalBreaks(): boolean
        isShowTextBoundaries(): boolean
        setDisplayBackgrounds(dbg: boolean): void
        setDisplayPageBoundaries(dpb: boolean): void
        setDisplaySmartTags(ds: boolean): void
        isShowFirstLineOnly(): boolean
        setShowFirstLineOnly(b: boolean): void
        setMailMergeDataView(mmdv: boolean): void
        setRevisionsBalloonSide(rbs: int): void
        getRevisionsBalloonSide(): int
        setRevisionsBalloonWidth(rbw: double): void
        getRevisionsBalloonWidth(): double
        setShowFieldCodes(sfc: boolean): void
        setShowFormatChanges(sfc: boolean): void
        setShowHiddenText(sht: boolean): void
        setShowInkAnnotations(sia: boolean): void
        setShowMainTextLayer(smtl: boolean): void
        setShowObjectAnchors(soa: boolean): void
        setShowOptionalBreaks(sob: boolean): void
        setShowParagraphs(sp: boolean): void
        setShowTextBoundaries(stb: boolean): void
        changeViewTypeToEIO(t: int): int
        changeViewTypeToMS(type: int): int
        getRevisionsFilter(): word.RevisionsFilter
        previousHeaderFooter(): void
        setDraft(df: boolean): void
        isReadingLayoutAllowMultiplePages(): boolean
        isRevisionsBalloonShowConnectingLines(): boolean
        setReadingLayoutAllowMultiplePages(rlamp: boolean): void
        setRevisionsBalloonShowConnectingLines(rbsc: boolean): void
        getMDoc(): any
        isReadingLayoutActualView(): boolean
        isShowInsertionsAndDeletions(): boolean
        isShowPicturePlaceHolders(): boolean
        setReadingLayoutActualView(rlav: boolean): void
        setRevisionsBalloonWidthType(rbwt: int): void
        getRevisionsBalloonWidthType(): int
        setShowInsertionsAndDeletions(sinsa: boolean): void
        setShowPicturePlaceHolders(spph: boolean): void
        isMagnifier(): boolean
        isShowAnimation(): boolean
        isShowDrawings(): boolean
        isShowFieldCodes(): boolean
        isShowHighlight(): boolean
        isShowHyphens(): boolean
        isShowXMLMarkup(): boolean
        isWrapToWindow(): boolean
        setFieldShading(ds: int): void
        getFieldShading(): int
        setMagnifier(mag: boolean): void
        setMarkupMode(mode: int): void
        getMarkupMode(): int
        setRevisionsMode(rm: int): void
        getRevisionsMode(): int
        setShowAnimation(sani: boolean): void
        setShowBookmarks(sbm: boolean): void
        isShowBookmarks(): boolean
        setShowComments(sc: boolean): void
        isShowComments(): boolean
        setShowDrawings(sd: boolean): void
        isShowHiddenText(): boolean
        setShowHighlight(shl: boolean): void
        setShowHyphens(shy: boolean): void
        isShowParagraphs(): boolean
        setShowSpaces(ss: boolean): void
        isShowSpaces(): boolean
        setShowTabs(st: boolean): void
        isShowTabs(): boolean
        setShowXMLMarkup(sxm: int): void
        setSplitSpecial(ssp: int): void
        getSplitSpecial(): int
        setWrapToWindow(wtw: boolean): void
        collapseOutline(range: any): void
        expandOutline(range: any): void
        nextHeaderFooter(): void
        showAllHeadings(): void
        setShowFormat(show: boolean): void
        isShowFormat(): boolean
        showHeading(level: int): void
    }
    export interface View {
        getType(): int
        setType(t: int): void
        setFullScreen(fs: boolean): void
        isFullScreen(): boolean
        getZoom(): word.Zoom
        setShowRevisionsAndComments(srac: boolean): void
        isShowRevisionsAndComments(): boolean
        setTableGridlines(tgl: boolean): void
        setShadeEditableRanges(ser: int): void
        getShadeEditableRanges(): int
        getRevisionsView(): int
        setRevisionsView(rv: int): void
        getReviewers(): word.Reviewers
        isReadingLayout(): boolean
        setReadingLayout(rl: boolean): void
        setSeekView(sv: int): void
        getSeekView(): int
        isTableGridlines(): boolean
        isDraft(): boolean
        isShowAll(): boolean
        setShowAll(sa: boolean): void
        isDisplayBackgrounds(): boolean
        isDisplayPageBoundaries(): boolean
        isDisplaySmartTags(): boolean
        isMailMergeDataView(): boolean
        isShowFormatChanges(): boolean
        isShowInkAnnotations(): boolean
        isShowMainTextLayer(): boolean
        isShowObjectAnchors(): boolean
        isShowOptionalBreaks(): boolean
        isShowTextBoundaries(): boolean
        setDisplayBackgrounds(dbg: boolean): void
        setDisplayPageBoundaries(dpb: boolean): void
        setDisplaySmartTags(ds: boolean): void
        isShowFirstLineOnly(): boolean
        setShowFirstLineOnly(b: boolean): void
        setMailMergeDataView(mmdv: boolean): void
        setRevisionsBalloonSide(rbs: int): void
        getRevisionsBalloonSide(): int
        setRevisionsBalloonWidth(rbw: double): void
        getRevisionsBalloonWidth(): double
        setShowFieldCodes(sfc: boolean): void
        setShowFormatChanges(sfc: boolean): void
        setShowHiddenText(sht: boolean): void
        setShowInkAnnotations(sia: boolean): void
        setShowMainTextLayer(smtl: boolean): void
        setShowObjectAnchors(soa: boolean): void
        setShowOptionalBreaks(sob: boolean): void
        setShowParagraphs(sp: boolean): void
        setShowTextBoundaries(stb: boolean): void
        changeViewTypeToEIO(t: int): int
        changeViewTypeToMS(type: int): int
        getRevisionsFilter(): word.RevisionsFilter
        previousHeaderFooter(): void
        setDraft(df: boolean): void
        isReadingLayoutAllowMultiplePages(): boolean
        isRevisionsBalloonShowConnectingLines(): boolean
        setReadingLayoutAllowMultiplePages(rlamp: boolean): void
        setRevisionsBalloonShowConnectingLines(rbsc: boolean): void
        getMDoc(): any
        isReadingLayoutActualView(): boolean
        isShowInsertionsAndDeletions(): boolean
        isShowPicturePlaceHolders(): boolean
        setReadingLayoutActualView(rlav: boolean): void
        setRevisionsBalloonWidthType(rbwt: int): void
        getRevisionsBalloonWidthType(): int
        setShowInsertionsAndDeletions(sinsa: boolean): void
        setShowPicturePlaceHolders(spph: boolean): void
        isMagnifier(): boolean
        isShowAnimation(): boolean
        isShowDrawings(): boolean
        isShowFieldCodes(): boolean
        isShowHighlight(): boolean
        isShowHyphens(): boolean
        isShowXMLMarkup(): boolean
        isWrapToWindow(): boolean
        setFieldShading(ds: int): void
        getFieldShading(): int
        setMagnifier(mag: boolean): void
        setMarkupMode(mode: int): void
        getMarkupMode(): int
        setRevisionsMode(rm: int): void
        getRevisionsMode(): int
        setShowAnimation(sani: boolean): void
        setShowBookmarks(sbm: boolean): void
        isShowBookmarks(): boolean
        setShowComments(sc: boolean): void
        isShowComments(): boolean
        setShowDrawings(sd: boolean): void
        isShowHiddenText(): boolean
        setShowHighlight(shl: boolean): void
        setShowHyphens(shy: boolean): void
        isShowParagraphs(): boolean
        setShowSpaces(ss: boolean): void
        isShowSpaces(): boolean
        setShowTabs(st: boolean): void
        isShowTabs(): boolean
        setShowXMLMarkup(sxm: int): void
        setSplitSpecial(ssp: int): void
        getSplitSpecial(): int
        setWrapToWindow(wtw: boolean): void
        collapseOutline(range: any): void
        expandOutline(range: any): void
        nextHeaderFooter(): void
        showAllHeadings(): void
        setShowFormat(show: boolean): void
        isShowFormat(): boolean
        showHeading(level: int): void
    }
    export interface WebOptions {
        getEncoding(): int
        isOptimizeForBrowser(): boolean
        setOptimizeForBrowser(ofb: boolean): void
        isOrganizeInFolder(): boolean
        setOrganizeInFolder(oif: boolean): void
        setUseLongFileNames(ufn: boolean): void
        isUseLongFileNames(): boolean
        getFolderSuffix(): string
        isAllowPNG(): boolean
        setAllowPNG(png: boolean): void
        getBrowserLevel(): int
        setBrowserLevel(bl: int): void
        setEncoding(end: int): void
        getPixelsPerInch(): int
        setPixelsPerInch(ppInch: int): void
        setRelyOnCSS(css: boolean): void
        isRelyOnVML(): boolean
        setRelyOnVML(rov: boolean): void
        getScreenSize(): int
        setScreenSize(sSize: int): void
        setTargetBrowser(tb: int): void
        getTargetBrowser(): int
        useDefaultFolderSuffix(): void
        getRelyOnCSS(): boolean
    }
    export interface WebOptions {
        getEncoding(): int
        isOptimizeForBrowser(): boolean
        setOptimizeForBrowser(ofb: boolean): void
        isOrganizeInFolder(): boolean
        setOrganizeInFolder(oif: boolean): void
        setUseLongFileNames(ufn: boolean): void
        isUseLongFileNames(): boolean
        getFolderSuffix(): string
        isAllowPNG(): boolean
        setAllowPNG(png: boolean): void
        getBrowserLevel(): int
        setBrowserLevel(bl: int): void
        setEncoding(end: int): void
        getPixelsPerInch(): int
        setPixelsPerInch(ppInch: int): void
        setRelyOnCSS(css: boolean): void
        isRelyOnVML(): boolean
        setRelyOnVML(rov: boolean): void
        getScreenSize(): int
        setScreenSize(sSize: int): void
        setTargetBrowser(tb: int): void
        getTargetBrowser(): int
        useDefaultFolderSuffix(): void
        getRelyOnCSS(): boolean
    }
    export interface Window {
        close(saveChanges: any, routeDocument: any): void
        getType(): int
        getIndex(): int
        isActive(): boolean
        getWidth(): int
        getLeft(): int
        setLeft(l: int): void
        setTop(in: int): void
        getTop(): int
        setWidth(w: int): void
        activate(): void
        setCaption(cap: string): void
        getCaption(): string
        isDisplayScreenTips(): boolean
        setDisplayScreenTips(dst: boolean): void
        getHeight(): int
        setHeight(h: int): void
        getSelection(): word.Selection
        getUsableHeight(): int
        getUsableWidth(): int
        isVisible(): boolean
        setVisible(v: boolean): void
        getWindowState(): int
        setWindowState(pWindowState: int): void
        newWindow(): word.Window
        printOut(background: any, append: any, range: any, outputFileName: any, from: any, to: any, item: any, copies: any, pages: string, pageType: any, printToFile: any, collate: any, activePrinterMacGX: any, manualDuplexPrint: any, printZoomColumn: any, printZoomRow: any, printZoomPaperWidth: any, printZoomPaperHeight: any): void
        getDocument(): word.Document
        getPrevious(): word.Window
        rangeFromPoint(x: int, y: int): any
        getNext(): word.Window
        getView(): word.View
        getActivePane(): word.Pane
        getVerticalPercentScrolled(): int
        setVerticalPercentScrolled(vps: int): void
        setHorizontalPercentScrolled(hps: int): void
        isDisplayVerticalRuler(): boolean
        setDisplayVerticalRuler(dvr: boolean): void
        isDisplayRulers(): boolean
        setDisplayRulers(dr: boolean): void
        largeScroll(down: any, up: any, toRight: any, toLeft: any): void
        pageScroll(down: any, up: any): void
        smallScroll(down: any, up: any, toRight: any, toLeft: any): void
        isDisplayLeftScrollBar(): boolean
        isDisplayRightRuler(): boolean
        isEnvelopeVisible(): boolean
        setDisplayLeftScrollBar(dls: boolean): void
        setDisplayRightRuler(drr: boolean): void
        setEnvelopeVisible(ev: boolean): void
        setStyleAreaWidth(saw: double): void
        getStyleAreaWidth(): double
        getContainSolidObject(mediator: any, modelPoint: any, objects: any[], mainControl: any): any
        getPanes(): word.Panes
        setSplit(sp: boolean): void
        isSplit(): boolean
        getPoint(obj: any): word.Rectangle
        setFocus(): void
        toggleShowAllReviewers(): void
        isDisplayHorizontalScrollBar(): boolean
        isDisplayVerticalScrollBar(): boolean
        isHorizontalPercentScrolled(): boolean
        setDisplayHorizontalScrollBar(dhs: boolean): void
        setDisplayVerticalScrollBar(dvs: boolean): void
        setDocumentMapPercentWidth(dmpw: int): void
        getDocumentMapPercentWidth(): int
        isDocumentMap(): boolean
        isThumbnails(): boolean
        getMWindow(): any
        setDocumentMap(dm: boolean): void
        setIMEMode(ime: int): void
        getIMEMode(): int
        setSplitVertical(spv: int): void
        getSplitVertical(): int
        setThumbnails(thum: boolean): void
        getWindowNumber(): int
        scrollIntoView(obj: any, start: any): void
        toggleRibbon(): void
    }
    export interface Window {
        close(saveChanges: any, routeDocument: any): void
        getType(): int
        getIndex(): int
        isActive(): boolean
        getWidth(): int
        getLeft(): int
        setLeft(l: int): void
        setTop(in: int): void
        getTop(): int
        setWidth(w: int): void
        activate(): void
        setCaption(cap: string): void
        getCaption(): string
        isDisplayScreenTips(): boolean
        setDisplayScreenTips(dst: boolean): void
        getHeight(): int
        setHeight(h: int): void
        getSelection(): word.Selection
        getUsableHeight(): int
        getUsableWidth(): int
        isVisible(): boolean
        setVisible(v: boolean): void
        getWindowState(): int
        setWindowState(pWindowState: int): void
        newWindow(): word.Window
        printOut(background: any, append: any, range: any, outputFileName: any, from: any, to: any, item: any, copies: any, pages: string, pageType: any, printToFile: any, collate: any, activePrinterMacGX: any, manualDuplexPrint: any, printZoomColumn: any, printZoomRow: any, printZoomPaperWidth: any, printZoomPaperHeight: any): void
        getDocument(): word.Document
        getPrevious(): word.Window
        rangeFromPoint(x: int, y: int): any
        getNext(): word.Window
        getView(): word.View
        getActivePane(): word.Pane
        getVerticalPercentScrolled(): int
        setVerticalPercentScrolled(vps: int): void
        setHorizontalPercentScrolled(hps: int): void
        isDisplayVerticalRuler(): boolean
        setDisplayVerticalRuler(dvr: boolean): void
        isDisplayRulers(): boolean
        setDisplayRulers(dr: boolean): void
        largeScroll(down: any, up: any, toRight: any, toLeft: any): void
        pageScroll(down: any, up: any): void
        smallScroll(down: any, up: any, toRight: any, toLeft: any): void
        isDisplayLeftScrollBar(): boolean
        isDisplayRightRuler(): boolean
        isEnvelopeVisible(): boolean
        setDisplayLeftScrollBar(dls: boolean): void
        setDisplayRightRuler(drr: boolean): void
        setEnvelopeVisible(ev: boolean): void
        setStyleAreaWidth(saw: double): void
        getStyleAreaWidth(): double
        getContainSolidObject(mediator: any, modelPoint: any, objects: any[], mainControl: any): any
        getPanes(): word.Panes
        setSplit(sp: boolean): void
        isSplit(): boolean
        getPoint(obj: any): word.Rectangle
        setFocus(): void
        toggleShowAllReviewers(): void
        isDisplayHorizontalScrollBar(): boolean
        isDisplayVerticalScrollBar(): boolean
        isHorizontalPercentScrolled(): boolean
        setDisplayHorizontalScrollBar(dhs: boolean): void
        setDisplayVerticalScrollBar(dvs: boolean): void
        setDocumentMapPercentWidth(dmpw: int): void
        getDocumentMapPercentWidth(): int
        isDocumentMap(): boolean
        isThumbnails(): boolean
        getMWindow(): any
        setDocumentMap(dm: boolean): void
        setIMEMode(ime: int): void
        getIMEMode(): int
        setSplitVertical(spv: int): void
        getSplitVertical(): int
        setThumbnails(thum: boolean): void
        getWindowNumber(): int
        scrollIntoView(obj: any, start: any): void
        toggleRibbon(): void
    }
    export interface Windows {
        add(win: any): word.Window
        item(index: any): word.Window
        getCount(): int
        getActiveWindow(): word.Window
        breakSideBySide(): boolean
        arrange(arrangeStyle: any): void
        compareSideBySideWith(doc: any): boolean
        resetPositionsSideBySide(): void
        setSyncScrollingSideBySide(pSyncScrollingSideBySide: boolean): void
        isSyncScrollingSideBySide(): boolean
    }
    export interface Windows {
        add(win: any): word.Window
        item(index: any): word.Window
        getCount(): int
        getActiveWindow(): word.Window
        breakSideBySide(): boolean
        arrange(arrangeStyle: any): void
        compareSideBySideWith(doc: any): boolean
        resetPositionsSideBySide(): void
        setSyncScrollingSideBySide(pSyncScrollingSideBySide: boolean): void
        isSyncScrollingSideBySide(): boolean
    }
    export interface WordBasic {
        abs(number: any): any
        activate(name: string): void
        getActiveDocument(): word.Document
        fileList(number: int): void
        fieldSeparator$(separator: string): void
        fileClose(): void
        fileCloseAll(): void
        fileNewDefault(): void
        fileName$(num: int): string
        fileName$(): string
        filePreview(): void
        filePrint(): void
        filePrintDefault(): void
        getApp(): word.Application
        fileExit(): void
        fileOpen(filePath: string): void
        isFileConfirmConversions(): boolean
        fileNameFromWindow$(): string
        rejectAllChangesShown(): void
        rejectAllChangesInDoc(): void
        deleteAllCommentsInDoc(): void
        setFileConfirmConversions(closePicture: boolean): void
    }
    export interface WordBasic {
        abs(number: any): any
        activate(name: string): void
        getActiveDocument(): word.Document
        fileList(number: int): void
        fieldSeparator$(separator: string): void
        fileClose(): void
        fileCloseAll(): void
        fileNewDefault(): void
        fileName$(num: int): string
        fileName$(): string
        filePreview(): void
        filePrint(): void
        filePrintDefault(): void
        getApp(): word.Application
        fileExit(): void
        fileOpen(filePath: string): void
        isFileConfirmConversions(): boolean
        fileNameFromWindow$(): string
        rejectAllChangesShown(): void
        rejectAllChangesInDoc(): void
        deleteAllCommentsInDoc(): void
        setFileConfirmConversions(closePicture: boolean): void
    }
    export interface Words {
        getFirst(): word.Range
        item(index: int): word.Range
        getLast(): word.Range
        getCount(): int
    }
    export interface Words {
        getFirst(): word.Range
        item(index: int): word.Range
        getLast(): word.Range
        getCount(): int
    }
    export interface WPUtil {
        getInstance(): word.WPUtil
        isProtected(doc: word.Document, start: long, end: long): boolean
        round(d1: double, n: int): double
        getDocument(parent: any): word.Document
        getIDMs2Yozo(msgCommand: string): number[]
        convertStyleIndexFormMs2Yozo(vMS: int): string
        getMSTextureIndexFromYozo(yozoIndex: int): int
        getYozoTextureIndexFromMS(msIndex: int): int
        mustSave(obj: office.IOfficeBase): void
    }
    export interface WPUtil {
        getInstance(): word.WPUtil
        isProtected(doc: word.Document, start: long, end: long): boolean
        round(d1: double, n: int): double
        getDocument(parent: any): word.Document
        getIDMs2Yozo(msgCommand: string): number[]
        convertStyleIndexFormMs2Yozo(vMS: int): string
        getMSTextureIndexFromYozo(yozoIndex: int): int
        getYozoTextureIndexFromMS(msIndex: int): int
        mustSave(obj: office.IOfficeBase): void
    }
    export interface WrapFormat {
        getType(): int
        setType(wt: int): void
        getAllowOverlap(): int
        setAllowOverlap(ao: int): void
        getDistanceLeft(): double
        setDistanceLeft(df: double): void
        getDistanceRight(): double
        setDistanceRight(dr: double): void
        getDistanceTop(): double
        setDistanceTop(dt: double): void
        getDistanceBottom(): double
        setDistanceBottom(db: double): void
        setSide(st: int): void
        getSide(): int
    }
    export interface WrapFormat {
        getType(): int
        setType(wt: int): void
        getAllowOverlap(): int
        setAllowOverlap(ao: int): void
        getDistanceLeft(): double
        setDistanceLeft(df: double): void
        getDistanceRight(): double
        setDistanceRight(dr: double): void
        getDistanceTop(): double
        setDistanceTop(dt: double): void
        getDistanceBottom(): double
        setDistanceBottom(db: double): void
        setSide(st: int): void
        getSide(): int
    }
    export interface XMLChildNodeSuggestion {
        insert(range: word.Range): word.XMLNode
        getBaseName(): string
        getNamespaceURI(): string
        getXMLSchemaReference(): word.XMLSchemaReference
    }
    export interface XMLChildNodeSuggestion {
        insert(range: word.Range): word.XMLNode
        getBaseName(): string
        getNamespaceURI(): string
        getXMLSchemaReference(): word.XMLSchemaReference
    }
    export interface XMLChildNodeSuggestions {
        item(index: int): word.XMLChildNodeSuggestion
        getCount(): int
    }
    export interface XMLChildNodeSuggestions {
        item(index: int): word.XMLChildNodeSuggestion
        getCount(): int
    }
    export interface XMLNamespace {
        getLocation(allUsers: boolean): string
        delete(): void
        setLocation(allUsers: boolean, loc: string): void
        getXSLTransforms(): word.XSLTransforms
        attachToDocument(doc: word.Document): void
        setDefaultTransform(allUsers: boolean, trans: word.XSLTransform): void
        getDefaultTransform(allUsers: boolean): word.XSLTransform
        setAlias(allUser: boolean, ali: string): void
        getAlias(allUser: boolean): string
        getURI(): string
    }
    export interface XMLNamespace {
        getLocation(allUsers: boolean): string
        delete(): void
        setLocation(allUsers: boolean, loc: string): void
        getXSLTransforms(): word.XSLTransforms
        attachToDocument(doc: word.Document): void
        setDefaultTransform(allUsers: boolean, trans: word.XSLTransform): void
        getDefaultTransform(allUsers: boolean): word.XSLTransform
        setAlias(allUser: boolean, ali: string): void
        getAlias(allUser: boolean): string
        getURI(): string
    }
    export interface XMLNamespaces {
        add(path: string, namespaceURI: any, alias: any, isInstallForAllUsers: boolean): word.XMLNamespace
        item(index: any): word.XMLNamespace
        getCount(): int
        installManifest(path: string, isInstallForAllUsers: boolean): void
    }
    export interface XMLNamespaces {
        add(path: string, namespaceURI: any, alias: any, isInstallForAllUsers: boolean): word.XMLNamespace
        item(index: any): word.XMLNamespace
        getCount(): int
        installManifest(path: string, isInstallForAllUsers: boolean): void
    }
    export interface XMLNode {
        delete(): void
        copy(): void
        validate(): void
        getText(): string
        cut(): void
        setText(str: string): void
        getRange(): word.Range
        getPlaceholderText(): string
        setPlaceholderText(str: string): void
        getChildNodeSuggestions(): word.XMLChildNodeSuggestions
        getXML(isDataOnly: boolean): string
        selectNodes(xPath: string, prefixMapping: string, isFastSearchSkippingTextNodes: boolean): word.XMLNodes
        selectSingleNode(xPath: string, prefixMapping: string, isFastSearchSkippingTextNodes: boolean): word.XMLNode
        getXMLNodes(): word.XMLNodes
        getFirstChild(): word.XMLNode
        getLastChild(): word.XMLNode
        getLevel(): int
        getBaseName(): string
        getNamespaceURI(): string
        getChildNodes(): word.XMLNodes
        isHasChildNodes(): boolean
        getNextSibling(): word.XMLNode
        getNodeType(): int
        setNodeValue(str: string): void
        getNodeValue(): string
        getOwnerDocument(): word.Document
        getParentNode(): word.XMLNode
        getSmartTag(): word.SmartTag
        removeChild(childElement: word.XMLNode): void
        getPreviousSibling(): word.XMLNode
        getValidationErrorText(isAdvanced: boolean): string
        getValidationStatu(): int
        setValidationError(status: int, errorText: string, isClearedAutomatically: boolean): void
    }
    export interface XMLNode {
        delete(): void
        copy(): void
        validate(): void
        getText(): string
        cut(): void
        setText(str: string): void
        getRange(): word.Range
        getPlaceholderText(): string
        setPlaceholderText(str: string): void
        getChildNodeSuggestions(): word.XMLChildNodeSuggestions
        getXML(isDataOnly: boolean): string
        selectNodes(xPath: string, prefixMapping: string, isFastSearchSkippingTextNodes: boolean): word.XMLNodes
        selectSingleNode(xPath: string, prefixMapping: string, isFastSearchSkippingTextNodes: boolean): word.XMLNode
        getXMLNodes(): word.XMLNodes
        getFirstChild(): word.XMLNode
        getLastChild(): word.XMLNode
        getLevel(): int
        getBaseName(): string
        getNamespaceURI(): string
        getChildNodes(): word.XMLNodes
        isHasChildNodes(): boolean
        getNextSibling(): word.XMLNode
        getNodeType(): int
        setNodeValue(str: string): void
        getNodeValue(): string
        getOwnerDocument(): word.Document
        getParentNode(): word.XMLNode
        getSmartTag(): word.SmartTag
        removeChild(childElement: word.XMLNode): void
        getPreviousSibling(): word.XMLNode
        getValidationErrorText(isAdvanced: boolean): string
        getValidationStatu(): int
        setValidationError(status: int, errorText: string, isClearedAutomatically: boolean): void
    }
    export interface XMLNodes {
        add(name: string, namespace: string, range: word.Range): word.XMLNode
        item(index: int): word.XMLNode
        getCount(): int
    }
    export interface XMLNodes {
        add(name: string, namespace: string, range: word.Range): word.XMLNode
        item(index: int): word.XMLNode
        getCount(): int
    }
    export interface XMLSchemaReference {
        getLocation(): string
        delete(): void
        reload(): void
        getNamespaceURI(): string
    }
    export interface XMLSchemaReference {
        getLocation(): string
        delete(): void
        reload(): void
        getNamespaceURI(): string
    }
    export interface XMLSchemaReferences {
        add(namespaceURI: any, alias: any, fileName: any, isInstallForAllUsers: boolean): word.XMLSchemaReference
        item(index: any): word.XMLSchemaReference
        validate(): void
        getCount(): int
        isAllowSaveAsXMLWithoutValidation(): boolean
        setAllowSaveAsXMLWithoutValidation(bool: boolean): void
        isAutomaticValidation(): boolean
        isHideValidationErrors(): boolean
        isIgnoreMixedContent(): boolean
        isShowPlaceholderText(): boolean
        setAutomaticValidation(bool: boolean): void
        setHideValidationErrors(bool: boolean): void
        setIgnoreMixedContent(bool: boolean): void
        setShowPlaceholderText(bool: boolean): void
    }
    export interface XMLSchemaReferences {
        add(namespaceURI: any, alias: any, fileName: any, isInstallForAllUsers: boolean): word.XMLSchemaReference
        item(index: any): word.XMLSchemaReference
        validate(): void
        getCount(): int
        isAllowSaveAsXMLWithoutValidation(): boolean
        setAllowSaveAsXMLWithoutValidation(bool: boolean): void
        isAutomaticValidation(): boolean
        isHideValidationErrors(): boolean
        isIgnoreMixedContent(): boolean
        isShowPlaceholderText(): boolean
        setAutomaticValidation(bool: boolean): void
        setHideValidationErrors(bool: boolean): void
        setIgnoreMixedContent(bool: boolean): void
        setShowPlaceholderText(bool: boolean): void
    }
    export interface XSLTransform {
        getLocation(allUsers: boolean): string
        delete(): void
        setLocation(allUsers: boolean, loc: string): void
        setAlias(isAllUsers: boolean, str: string): void
        getAlias(isAllUsers: boolean): string
    }
    export interface XSLTransform {
        getLocation(allUsers: boolean): string
        delete(): void
        setLocation(allUsers: boolean, loc: string): void
        setAlias(isAllUsers: boolean, str: string): void
        getAlias(isAllUsers: boolean): string
    }
    export interface XSLTransforms {
        add(location: string, alias: any, isInstallForAllUsers: boolean): word.XSLTransform
        item(index: any): word.XSLTransform
        getCount(): int
    }
    export interface XSLTransforms {
        add(location: string, alias: any, isInstallForAllUsers: boolean): word.XSLTransform
        item(index: any): word.XSLTransform
        getCount(): int
    }
    export interface Zoom {
        getPercentage(): int
        setPageColumns(pc: int): void
        getPageRows(): int
        getPageColumns(): int
        setPageFit(pf: int): void
        getPageFit(): int
        setPageRows(pr: int): void
        setPercentage(zoom: int): void
    }
    export interface Zoom {
        getPercentage(): int
        setPageColumns(pc: int): void
        getPageRows(): int
        getPageColumns(): int
        setPageFit(pf: int): void
        getPageFit(): int
        setPageRows(pr: int): void
        setPercentage(zoom: int): void
    }
    export interface Zooms {
        item(viewType: int): word.Zoom
    }
    export interface Zooms {
        item(viewType: int): word.Zoom
    }
}
declare namespace word.constants {
    export const enum YwAlertLevel {
        ywAlertsAll = -1,
        ywAlertsMessageBox = -2,
        ywAlertsNone = 0

    }
    export const enum YwAlertLevel {
        ywAlertsAll = -1,
        ywAlertsMessageBox = -2,
        ywAlertsNone = 0

    }
    export const enum YwArabicNumeral {
        ywNumeralArabic = 0,
        ywNumeralHindi = 1,
        ywNumeralContext = 2,
        ywNumeralSystem = 3

    }
    export const enum YwArabicNumeral {
        ywNumeralArabic = 0,
        ywNumeralHindi = 1,
        ywNumeralContext = 2,
        ywNumeralSystem = 3

    }
    export const enum YwAraSpeller {
        ywBoth = 3,
        ywFinalYaa = 2,
        ywInitialAlef = 1,
        ywNone = 0

    }
    export const enum YwAraSpeller {
        ywBoth = 3,
        ywFinalYaa = 2,
        ywInitialAlef = 1,
        ywNone = 0

    }
    export const enum YwArrangeStyle {
        ywIcons = 1,
        ywTiled = 0

    }
    export const enum YwArrangeStyle {
        ywIcons = 1,
        ywTiled = 0

    }
    export const enum YwAutoFitBehavior {
        ywAutoFitContent = 1,
        ywAutoFitFixed = 0,
        ywAutoFitWindow = 2

    }
    export const enum YwAutoFitBehavior {
        ywAutoFitContent = 1,
        ywAutoFitFixed = 0,
        ywAutoFitWindow = 2

    }
    export const enum YwBaselineAlignment {
        ywBaselineAlignAuto = 4,
        ywBaselineAlignBaseline = 2,
        ywBaselineAlignCenter = 1,
        ywBaselineAlignFarEast50 = 3,
        ywBaselineAlignTop = 0

    }
    export const enum YwBaselineAlignment {
        ywBaselineAlignAuto = 4,
        ywBaselineAlignBaseline = 2,
        ywBaselineAlignCenter = 1,
        ywBaselineAlignFarEast50 = 3,
        ywBaselineAlignTop = 0

    }
    export const enum YwBookmarkSortBy {
        ywSortByLocation = 1,
        ywSortByName = 0

    }
    export const enum YwBookmarkSortBy {
        ywSortByLocation = 1,
        ywSortByName = 0

    }
    export const enum YwBorderDistanceFrom {
        ywBorderDistanceFromPageEdge = 1,
        ywBorderDistanceFromText = 0

    }
    export const enum YwBorderDistanceFrom {
        ywBorderDistanceFromPageEdge = 1,
        ywBorderDistanceFromText = 0

    }
    export const enum YwBorderType {
        ywBorderBottom = -3,
        ywBorderDiagonalDown = -7,
        ywBorderDiagonalUp = -8,
        ywBorderHorizontal = -5,
        ywBorderLeft = -2,
        ywBorderRight = -4,
        ywBorderTop = -1,
        ywBorderVertical = -6

    }
    export const enum YwBorderType {
        ywBorderBottom = -3,
        ywBorderDiagonalDown = -7,
        ywBorderDiagonalUp = -8,
        ywBorderHorizontal = -5,
        ywBorderLeft = -2,
        ywBorderRight = -4,
        ywBorderTop = -1,
        ywBorderVertical = -6

    }
    export const enum YwBreakType {
        ywPageBreak = 7,
        ywColumnBreak = 8,
        ywLineBreak = 6,
        ywSectionBreakNextPage = 2,
        ywSectionBreakContinuous = 3,
        ywSectionBreakEvenPage = 4,
        ywSectionBreakOddPage = 5

    }
    export const enum YwBreakType {
        ywPageBreak = 7,
        ywColumnBreak = 8,
        ywLineBreak = 6,
        ywSectionBreakNextPage = 2,
        ywSectionBreakContinuous = 3,
        ywSectionBreakEvenPage = 4,
        ywSectionBreakOddPage = 5

    }
    export const enum YwBrowseTarget {
        ywBrowseComment = 3,
        ywBrowseEdit = 10,
        ywBrowseEndnote = 5,
        ywBrowseField = 6,
        ywBrowseFind = 11,
        ywBrowseFootnote = 4,
        ywBrowseGoTo = 12,
        ywBrowseGraphic = 8,
        ywBrowseHeading = 9,
        ywBrowsePage = 1,
        ywBrowseSection = 2,
        ywBrowseTable = 7

    }
    export const enum YwBrowseTarget {
        ywBrowseComment = 3,
        ywBrowseEdit = 10,
        ywBrowseEndnote = 5,
        ywBrowseField = 6,
        ywBrowseFind = 11,
        ywBrowseFootnote = 4,
        ywBrowseGoTo = 12,
        ywBrowseGraphic = 8,
        ywBrowseHeading = 9,
        ywBrowsePage = 1,
        ywBrowseSection = 2,
        ywBrowseTable = 7

    }
    export const enum YwBuiltinStyle {
        ywStyleBlockQuotation = -85,
        ywStyleBodyText = -67,
        ywStyleBodyText2 = -81,
        ywStyleBodyText3 = -82,
        ywStyleBodyTextFirstIndent = -78,
        ywStyleBodyTextFirstIndent2 = -79,
        ywStyleBodyTextIndent = -68,
        ywStyleBodyTextIndent2 = -83,
        ywStyleBodyTextIndent3 = -84,
        ywStyleBookTitle = -265,
        ywStyleCaption = -35,
        ywStyleClosing = -64,
        ywStyleCommentReference = -40,
        ywStyleCommentText = -31,
        ywStyleDate = -77,
        ywStyleDefaultParagraphFont = -66,
        ywStyleEmphasis = -89,
        ywStyleEndnoteReference = -43,
        ywStyleEndnoteText = -44,
        ywStyleEnvelopeAddress = -37,
        ywStyleEnvelopeReturn = -38,
        ywStyleFooter = -33,
        ywStyleFootnoteReference = -39,
        ywStyleFootnoteText = -30,
        ywStyleHeader = -32,
        ywStyleHeading1 = -2,
        ywStyleHeading2 = -3,
        ywStyleHeading3 = -4,
        ywStyleHeading4 = -5,
        ywStyleHeading5 = -6,
        ywStyleHeading6 = -7,
        ywStyleHeading7 = -8,
        ywStyleHeading8 = -9,
        ywStyleHeading9 = -10,
        ywStyleHtmlAcronym = -96,
        ywStyleHtmlAddress = -97,
        ywStyleHtmlCite = -98,
        ywStyleHtmlCode = -99,
        ywStyleHtmlDfn = -100,
        ywStyleHtmlKbd = -101,
        ywStyleHtmlNormal = -95,
        ywStyleHtmlPre = -102,
        ywStyleHtmlSamp = -103,
        ywStyleHtmlTt = -104,
        ywStyleHtmlVar = -105,
        ywStyleHyperlink = -86,
        ywStyleHyperlinkFollowed = -87,
        ywStyleIndex1 = -11,
        ywStyleIndex2 = -12,
        ywStyleIndex3 = -13,
        ywStyleIndex4 = -14,
        ywStyleIndex5 = -15,
        ywStyleIndex6 = -16,
        ywStyleIndex7 = -17,
        ywStyleIndex8 = -18,
        ywStyleIndex9 = -19,
        ywStyleIndexHeading = -34,
        ywStyleIntenseEmphasis = -262,
        ywStyleIntenseQuote = -182,
        ywStyleIntenseReference = -264,
        ywStyleLineNumber = -41,
        ywStyleList = -48,
        ywStyleList2 = -51,
        ywStyleList3 = -52,
        ywStyleList4 = -53,
        ywStyleList5 = -54,
        ywStyleListBullet = -49,
        ywStyleListBullet2 = -55,
        ywStyleListBullet3 = -56,
        ywStyleListBullet4 = -57,
        ywStyleListBullet5 = -58,
        ywStyleListContinue = -69,
        ywStyleListContinue2 = -70,
        ywStyleListContinue3 = -71,
        ywStyleListContinue4 = -72,
        ywStyleListContinue5 = -73,
        ywStyleListNumber = -50,
        ywStyleListNumber2 = -59,
        ywStyleListNumber3 = -60,
        ywStyleListNumber4 = -61,
        ywStyleListNumber5 = -62,
        ywStyleListParagraph = -180,
        ywStyleMacroText = -46,
        ywStyleMessageHeader = -74,
        ywStyleNavPane = -90,
        ywStyleNormal = -1,
        ywStyleNormalIndent = -29,
        ywStyleNormalObject = -158,
        ywStyleNormalTable = -106,
        ywStyleNoteHeading = -80,
        ywStylePageNumber = -42,
        ywStylePlainText = -91,
        ywStyleQuote = -181,
        ywStyleSalutation = -76,
        ywStyleSignature = -65,
        ywStyleStrong = -88,
        ywStyleSubtitle = -75,
        ywdStyleSubtleEmphasis = -261,
        ywStyleSubtleReference = -263,
        ywStyleTableColorfulGrid = -172,
        ywStyleTableColorfulList = -171,
        ywStyleTableColorfulShading = -170,
        ywStyleTableDarkList = -169,
        ywStyleTableLightGrid = -161,
        ywStyleTableLightGridAccent1 = -175,
        ywStyleTableLightList = -160,
        ywStyleTableLightListAccent1 = -174,
        ywStyleTableLightShading = -159,
        ywStyleTableLightShadingAccent1 = -173,
        ywStyleTableMediumGrid1 = -166,
        ywStyleTableMediumGrid2 = -167,
        ywStyleTableMediumGrid3 = -168,
        ywStyleTableMediumList1 = -164,
        ywStyleTableMediumList1Accent1 = -178,
        ywStyleTableMediumList2 = -165,
        ywStyleTableMediumShading1 = -162,
        ywStyleTableMediumShading1Accent1 = -176,
        ywStyleTableMediumShading2 = -163,
        ywStyleTableMediumShading2Accent1 = -177,
        ywStyleTableOfAuthorities = -45,
        ywStyleTableOfFigures = -36,
        ywStyleTitle = -63,
        ywStyleTOAHeading = -47,
        ywStyleTOC1 = -20,
        ywStyleTOC2 = -21,
        ywStyleTOC3 = -22,
        ywStyleTOC4 = -23,
        ywStyleTOC5 = -24,
        ywStyleTOC6 = -25,
        ywStyleTOC7 = -26,
        ywStyleTOC8 = -27,
        ywStyleTOC9 = -28

    }
    export const enum YwBuiltinStyle {
        ywStyleBlockQuotation = -85,
        ywStyleBodyText = -67,
        ywStyleBodyText2 = -81,
        ywStyleBodyText3 = -82,
        ywStyleBodyTextFirstIndent = -78,
        ywStyleBodyTextFirstIndent2 = -79,
        ywStyleBodyTextIndent = -68,
        ywStyleBodyTextIndent2 = -83,
        ywStyleBodyTextIndent3 = -84,
        ywStyleBookTitle = -265,
        ywStyleCaption = -35,
        ywStyleClosing = -64,
        ywStyleCommentReference = -40,
        ywStyleCommentText = -31,
        ywStyleDate = -77,
        ywStyleDefaultParagraphFont = -66,
        ywStyleEmphasis = -89,
        ywStyleEndnoteReference = -43,
        ywStyleEndnoteText = -44,
        ywStyleEnvelopeAddress = -37,
        ywStyleEnvelopeReturn = -38,
        ywStyleFooter = -33,
        ywStyleFootnoteReference = -39,
        ywStyleFootnoteText = -30,
        ywStyleHeader = -32,
        ywStyleHeading1 = -2,
        ywStyleHeading2 = -3,
        ywStyleHeading3 = -4,
        ywStyleHeading4 = -5,
        ywStyleHeading5 = -6,
        ywStyleHeading6 = -7,
        ywStyleHeading7 = -8,
        ywStyleHeading8 = -9,
        ywStyleHeading9 = -10,
        ywStyleHtmlAcronym = -96,
        ywStyleHtmlAddress = -97,
        ywStyleHtmlCite = -98,
        ywStyleHtmlCode = -99,
        ywStyleHtmlDfn = -100,
        ywStyleHtmlKbd = -101,
        ywStyleHtmlNormal = -95,
        ywStyleHtmlPre = -102,
        ywStyleHtmlSamp = -103,
        ywStyleHtmlTt = -104,
        ywStyleHtmlVar = -105,
        ywStyleHyperlink = -86,
        ywStyleHyperlinkFollowed = -87,
        ywStyleIndex1 = -11,
        ywStyleIndex2 = -12,
        ywStyleIndex3 = -13,
        ywStyleIndex4 = -14,
        ywStyleIndex5 = -15,
        ywStyleIndex6 = -16,
        ywStyleIndex7 = -17,
        ywStyleIndex8 = -18,
        ywStyleIndex9 = -19,
        ywStyleIndexHeading = -34,
        ywStyleIntenseEmphasis = -262,
        ywStyleIntenseQuote = -182,
        ywStyleIntenseReference = -264,
        ywStyleLineNumber = -41,
        ywStyleList = -48,
        ywStyleList2 = -51,
        ywStyleList3 = -52,
        ywStyleList4 = -53,
        ywStyleList5 = -54,
        ywStyleListBullet = -49,
        ywStyleListBullet2 = -55,
        ywStyleListBullet3 = -56,
        ywStyleListBullet4 = -57,
        ywStyleListBullet5 = -58,
        ywStyleListContinue = -69,
        ywStyleListContinue2 = -70,
        ywStyleListContinue3 = -71,
        ywStyleListContinue4 = -72,
        ywStyleListContinue5 = -73,
        ywStyleListNumber = -50,
        ywStyleListNumber2 = -59,
        ywStyleListNumber3 = -60,
        ywStyleListNumber4 = -61,
        ywStyleListNumber5 = -62,
        ywStyleListParagraph = -180,
        ywStyleMacroText = -46,
        ywStyleMessageHeader = -74,
        ywStyleNavPane = -90,
        ywStyleNormal = -1,
        ywStyleNormalIndent = -29,
        ywStyleNormalObject = -158,
        ywStyleNormalTable = -106,
        ywStyleNoteHeading = -80,
        ywStylePageNumber = -42,
        ywStylePlainText = -91,
        ywStyleQuote = -181,
        ywStyleSalutation = -76,
        ywStyleSignature = -65,
        ywStyleStrong = -88,
        ywStyleSubtitle = -75,
        ywdStyleSubtleEmphasis = -261,
        ywStyleSubtleReference = -263,
        ywStyleTableColorfulGrid = -172,
        ywStyleTableColorfulList = -171,
        ywStyleTableColorfulShading = -170,
        ywStyleTableDarkList = -169,
        ywStyleTableLightGrid = -161,
        ywStyleTableLightGridAccent1 = -175,
        ywStyleTableLightList = -160,
        ywStyleTableLightListAccent1 = -174,
        ywStyleTableLightShading = -159,
        ywStyleTableLightShadingAccent1 = -173,
        ywStyleTableMediumGrid1 = -166,
        ywStyleTableMediumGrid2 = -167,
        ywStyleTableMediumGrid3 = -168,
        ywStyleTableMediumList1 = -164,
        ywStyleTableMediumList1Accent1 = -178,
        ywStyleTableMediumList2 = -165,
        ywStyleTableMediumShading1 = -162,
        ywStyleTableMediumShading1Accent1 = -176,
        ywStyleTableMediumShading2 = -163,
        ywStyleTableMediumShading2Accent1 = -177,
        ywStyleTableOfAuthorities = -45,
        ywStyleTableOfFigures = -36,
        ywStyleTitle = -63,
        ywStyleTOAHeading = -47,
        ywStyleTOC1 = -20,
        ywStyleTOC2 = -21,
        ywStyleTOC3 = -22,
        ywStyleTOC4 = -23,
        ywStyleTOC5 = -24,
        ywStyleTOC6 = -25,
        ywStyleTOC7 = -26,
        ywStyleTOC8 = -27,
        ywStyleTOC9 = -28

    }
    export const enum YwCaptionLabelID {
        ywCaptionEquation = -3,
        ywCaptionFigure = -1,
        ywCaptionTable = -2

    }
    export const enum YwCaptionLabelID {
        ywCaptionEquation = -3,
        ywCaptionFigure = -1,
        ywCaptionTable = -2

    }
    export const enum YwCaptionPosition {
        ywCaptionPositionAbove = 0,
        ywCaptionPositionBelow = 1

    }
    export const enum YwCaptionPosition {
        ywCaptionPositionAbove = 0,
        ywCaptionPositionBelow = 1

    }
    export const enum YwCellVerticalAlignment {
        YwCellAlignVerticalBottom = 3,
        YwCellAlignVerticalCenter = 1,
        YwCellAlignVerticalTop = 0

    }
    export const enum YwCellVerticalAlignment {
        YwCellAlignVerticalBottom = 3,
        YwCellAlignVerticalCenter = 1,
        YwCellAlignVerticalTop = 0

    }
    export const enum YwCharacterCase {
        ywFullWidth = 7,
        ywHalfWidth = 6,
        ywHiragana = 9,
        ywKatakana = 8,
        ywLowerCase = 0,
        ywNextCase = -1,
        ywTitleSentence = 4,
        ywTitleWord = 2,
        ywToggleCase = 5,
        ywUpperCase = 1

    }
    export const enum YwCharacterCase {
        ywFullWidth = 7,
        ywHalfWidth = 6,
        ywHiragana = 9,
        ywKatakana = 8,
        ywLowerCase = 0,
        ywNextCase = -1,
        ywTitleSentence = 4,
        ywTitleWord = 2,
        ywToggleCase = 5,
        ywUpperCase = 1

    }
    export const enum YwCharacterWidth {
        ywWidthFullWidth = 7,
        ywWidthHalfWidth = 6

    }
    export const enum YwCharacterWidth {
        ywWidthFullWidth = 7,
        ywWidthHalfWidth = 6

    }
    export const enum YwChevronConvertRule {
        ywAlwaysConvert = 1,
        ywNeverConvert = 0,
        ywAskToConvert = 3,
        ywAskToNotConvert = 2

    }
    export const enum YwChevronConvertRule {
        ywAlwaysConvert = 1,
        ywNeverConvert = 0,
        ywAskToConvert = 3,
        ywAskToNotConvert = 2

    }
    export const enum YwCollapseDirection {
        ywCollapseEnd = 0,
        ywCollapseStart = 1

    }
    export const enum YwCollapseDirection {
        ywCollapseEnd = 0,
        ywCollapseStart = 1

    }
    export interface YwColor {
        getTypeFromColor(color: any): int
        getColorFromType(colorType: int): any
    }
    export interface YwColor {
        getTypeFromColor(color: any): int
        getColorFromType(colorType: int): any
    }
    export const enum YwColorIndex {
        ywByAuthor = -1,
        ywNoHighlight = 0,
        ywAuto = 0,
        ywBlack = 1,
        ywBlue = 2,
        ywTurquoise = 3,
        ywBrightGreen = 4,
        ywPink = 5,
        ywRed = 6,
        ywYellow = 7,
        ywWhite = 8,
        ywDarkBlue = 9,
        ywTeal = 10,
        ywGreen = 11,
        ywViolet = 12,
        ywDarkRed = 13,
        ywDarkYellow = 14,
        ywGray50 = 15,
        ywGray25 = 16

    }
    export const enum YwColorIndex {
        ywByAuthor = -1,
        ywNoHighlight = 0,
        ywAuto = 0,
        ywBlack = 1,
        ywBlue = 2,
        ywTurquoise = 3,
        ywBrightGreen = 4,
        ywPink = 5,
        ywRed = 6,
        ywYellow = 7,
        ywWhite = 8,
        ywDarkBlue = 9,
        ywTeal = 10,
        ywGreen = 11,
        ywViolet = 12,
        ywDarkRed = 13,
        ywDarkYellow = 14,
        ywGray50 = 15,
        ywGray25 = 16

    }
    export const enum YwCompareDestination {

    }
    export const enum YwCompareDestination {

    }
    export const enum YwCompareTarget {
        ywCompareTargetCurrent = 1,
        ywCompareTargetNew = 2,
        ywCompareTargetSelected = 0

    }
    export const enum YwCompareTarget {
        ywCompareTargetCurrent = 1,
        ywCompareTargetNew = 2,
        ywCompareTargetSelected = 0

    }
    export const enum YwConditionCode {
        ywEvenColumnBanding = 7,
        ywEvenRowBanding = 3,
        ywFirstColumn = 4,
        ywFirstRow = 0,
        ywLastColumn = 5,
        ywLastRow = 1,
        ywNECell = 8,
        ywNWCell = 9,
        ywOddColumnBanding = 6,
        ywOddRowBanding = 2,
        ywSECell = 10,
        ywSWCell = 11

    }
    export const enum YwConditionCode {
        ywEvenColumnBanding = 7,
        ywEvenRowBanding = 3,
        ywFirstColumn = 4,
        ywFirstRow = 0,
        ywLastColumn = 5,
        ywLastRow = 1,
        ywNECell = 8,
        ywNWCell = 9,
        ywOddColumnBanding = 6,
        ywOddRowBanding = 2,
        ywSECell = 10,
        ywSWCell = 11

    }
    export const enum YwConstants {
        ywAutoPosition = 0,
        ywBackward = -1073741823,
        ywCreatorCode = 1297307460,
        ywFirst = 1,
        ywToggle = 9999998,
        ywUndefined = 9999999,
        ywForward = 1073741823

    }
    export const enum YwConstants {
        ywAutoPosition = 0,
        ywBackward = -1073741823,
        ywCreatorCode = 1297307460,
        ywFirst = 1,
        ywToggle = 9999998,
        ywUndefined = 9999999,
        ywForward = 1073741823

    }
    export const enum YwContentControlType {
        ywContentControlRichText = 0,
        ywContentControlText = 1,
        ywContentControlPicture = 2,
        ywContentControlComboBox = 3,
        ywContentControlDropdownList = 4,
        ywContentControlBuildingBlockGallery = 5,
        ywContentControlDate = 6,
        ywContentControlGroup = 7,
        ywContentControlCheckBox = 8,
        ywContentControlRepeatingSection = 9

    }
    export const enum YwContentControlType {
        ywContentControlRichText = 0,
        ywContentControlText = 1,
        ywContentControlPicture = 2,
        ywContentControlComboBox = 3,
        ywContentControlDropdownList = 4,
        ywContentControlBuildingBlockGallery = 5,
        ywContentControlDate = 6,
        ywContentControlGroup = 7,
        ywContentControlCheckBox = 8,
        ywContentControlRepeatingSection = 9

    }
    export const enum YwContinue {
        ywContinueDisabled = 0,
        ywResetList = 1,
        ywContinueList = 2

    }
    export const enum YwContinue {
        ywContinueDisabled = 0,
        ywResetList = 1,
        ywContinueList = 2

    }
    export const enum YwCountry {
        ywArgentina = 54,
        ywBrazil = 55,
        ywCanada = 2,
        ywChile = 56,
        ywChina = 86,
        ywDenmark = 45,
        ywFinland = 358,
        ywFrance = 33,
        ywGermany = 49,
        ywIceland = 354,
        ywItaly = 39,
        ywJapan = 81,
        ywKorea = 82,
        ywLatinAmerica = 3,
        ywMexico = 52,
        ywNetherlands = 31,
        ywNorway = 47,
        ywPeru = 51,
        ywSpain = 34,
        ywSweden = 46,
        ywTaiwan = 886,
        ywUK = 44,
        ywUS = 1,
        ywVenezuela = 58

    }
    export const enum YwCountry {
        ywArgentina = 54,
        ywBrazil = 55,
        ywCanada = 2,
        ywChile = 56,
        ywChina = 86,
        ywDenmark = 45,
        ywFinland = 358,
        ywFrance = 33,
        ywGermany = 49,
        ywIceland = 354,
        ywItaly = 39,
        ywJapan = 81,
        ywKorea = 82,
        ywLatinAmerica = 3,
        ywMexico = 52,
        ywNetherlands = 31,
        ywNorway = 47,
        ywPeru = 51,
        ywSpain = 34,
        ywSweden = 46,
        ywTaiwan = 886,
        ywUK = 44,
        ywUS = 1,
        ywVenezuela = 58

    }
    export const enum YwCursorMovement {
        ywCursorMovementLogical = 0,
        ywCursorMovementVisual = 1

    }
    export const enum YwCursorMovement {
        ywCursorMovementLogical = 0,
        ywCursorMovementVisual = 1

    }
    export const enum YwCursorType {
        ywCursorIBeam = 1,
        ywCursorNormal = 2,
        ywCursorNorthwestArrow = 3,
        ywCursorWait = 0

    }
    export const enum YwCursorType {
        ywCursorIBeam = 1,
        ywCursorNormal = 2,
        ywCursorNorthwestArrow = 3,
        ywCursorWait = 0

    }
    export const enum YwCustomLabelPageSize {
        wdCustomLabelA4 = 2,
        wdCustomLabelA4LS = 3,
        wdCustomLabelA5 = 4,
        wdCustomLabelA5LS = 5,
        wdCustomLabelB4JIS = 13,
        wdCustomLabelB5 = 6,
        wdCustomLabelFanfold = 8,
        wdCustomLabelHigaki = 11,
        wdCustomLabelHigakiLS = 12,
        wdCustomLabelLetter = 0,
        wdCustomLabelLetterLS = 1,
        wdCustomLabelMini = 7,
        wdCustomLabelVertHalfSheet = 9,
        wdCustomLabelVertHalfSheetLS = 10

    }
    export const enum YwCustomLabelPageSize {
        wdCustomLabelA4 = 2,
        wdCustomLabelA4LS = 3,
        wdCustomLabelA5 = 4,
        wdCustomLabelA5LS = 5,
        wdCustomLabelB4JIS = 13,
        wdCustomLabelB5 = 6,
        wdCustomLabelFanfold = 8,
        wdCustomLabelHigaki = 11,
        wdCustomLabelHigakiLS = 12,
        wdCustomLabelLetter = 0,
        wdCustomLabelLetterLS = 1,
        wdCustomLabelMini = 7,
        wdCustomLabelVertHalfSheet = 9,
        wdCustomLabelVertHalfSheetLS = 10

    }
    export const enum YwDefaultFilePath {
        ywBorderArtPath = 19,
        ywStyleGalleryPath = 15,
        ywCurrentFolderPath = 14,
        ywTempFilePath = 13,
        ywProofingToolsPath = 12,
        ywTextConvertersPath = 11,
        ywGraphicsFiltersPath = 10,
        ywProgramPath = 9,
        ywStartupPath = 8,
        ywTutorialPath = 7,
        ywToolsPath = 6,
        ywAutoRecoverPath = 5,
        ywUserOptionsPath = 4,
        ywWorkgroupTemplatesPath = 3,
        ywUserTemplatesPath = 2,
        ywPicturesPath = 1,
        ywDocumentsPath = 0

    }
    export const enum YwDefaultFilePath {
        ywBorderArtPath = 19,
        ywStyleGalleryPath = 15,
        ywCurrentFolderPath = 14,
        ywTempFilePath = 13,
        ywProofingToolsPath = 12,
        ywTextConvertersPath = 11,
        ywGraphicsFiltersPath = 10,
        ywProgramPath = 9,
        ywStartupPath = 8,
        ywTutorialPath = 7,
        ywToolsPath = 6,
        ywAutoRecoverPath = 5,
        ywUserOptionsPath = 4,
        ywWorkgroupTemplatesPath = 3,
        ywUserTemplatesPath = 2,
        ywPicturesPath = 1,
        ywDocumentsPath = 0

    }
    export const enum YwDefaultListBehavior {
        ywWord8ListBehavior = 0,
        ywWord9ListBehavior = 1,
        ywWord10ListBehavior = 2

    }
    export const enum YwDefaultListBehavior {
        ywWord8ListBehavior = 0,
        ywWord9ListBehavior = 1,
        ywWord10ListBehavior = 2

    }
    export const enum YwDefaultTableBehavior {
        wdWord8TableBehavior = 0,
        wdWord9TableBehavior = 1

    }
    export const enum YwDefaultTableBehavior {
        wdWord8TableBehavior = 0,
        wdWord9TableBehavior = 1

    }
    export const enum YwDeleteCells {
        ywDeleteCellsEntireColumn = 3,
        ywDeleteCellsEntireRow = 2,
        ywDeleteCellsShiftLeft = 0,
        ywDeleteCellsShiftUp = 1

    }
    export const enum YwDeleteCells {
        ywDeleteCellsEntireColumn = 3,
        ywDeleteCellsEntireRow = 2,
        ywDeleteCellsShiftLeft = 0,
        ywDeleteCellsShiftUp = 1

    }
    export const enum YwDeletedTextMark {
        ywDeletedTextMarkDoubleStrikeThrough = 10,
        ywDeletedTextMarkColorOnly = 9,
        ywDeletedTextMarkDoubleUnderline = 8,
        ywDeletedTextMarkUnderline = 7,
        ywDeletedTextMarkItalic = 6,
        ywDeletedTextMarkBold = 5,
        ywDeletedTextMarkNone = 4,
        ywDeletedTextMarkPound = 3,
        ywDeletedTextMarkCaret = 2,
        ywDeletedTextMarkStrikeThrough = 1,
        ywDeletedTextMarkHidden = 0

    }
    export const enum YwDeletedTextMark {
        ywDeletedTextMarkDoubleStrikeThrough = 10,
        ywDeletedTextMarkColorOnly = 9,
        ywDeletedTextMarkDoubleUnderline = 8,
        ywDeletedTextMarkUnderline = 7,
        ywDeletedTextMarkItalic = 6,
        ywDeletedTextMarkBold = 5,
        ywDeletedTextMarkNone = 4,
        ywDeletedTextMarkPound = 3,
        ywDeletedTextMarkCaret = 2,
        ywDeletedTextMarkStrikeThrough = 1,
        ywDeletedTextMarkHidden = 0

    }
    export const enum YwDictionaryType {
        ywGrammar = 1,
        ywHangulHanjaConversion = 8,
        ywHangulHanjaConversionCustom = 9,
        ywHyphenation = 3,
        ywSpelling = 0,
        ywSpellingComplete = 4,
        ywSpellingCustom = 5,
        ywSpellingLegal = 6,
        ywSpellingMedical = 7,
        ywThesaurus = 2

    }
    export const enum YwDictionaryType {
        ywGrammar = 1,
        ywHangulHanjaConversion = 8,
        ywHangulHanjaConversionCustom = 9,
        ywHyphenation = 3,
        ywSpelling = 0,
        ywSpellingComplete = 4,
        ywSpellingCustom = 5,
        ywSpellingLegal = 6,
        ywSpellingMedical = 7,
        ywThesaurus = 2

    }
    export const enum YwDisableFeaturesIntroducedAfter {
        yw70 = 0,
        yw70FE = 1,
        yw80 = 2

    }
    export const enum YwDisableFeaturesIntroducedAfter {
        yw70 = 0,
        yw70FE = 1,
        yw80 = 2

    }
    export const enum YwDocumentKind {
        ywDocumentEmail = 2,
        ywDocumentLetter = 1,
        ywDocumentNotSpecified = 0

    }
    export const enum YwDocumentKind {
        ywDocumentEmail = 2,
        ywDocumentLetter = 1,
        ywDocumentNotSpecified = 0

    }
    export const enum YwDocumentType {
        ywTypeDocument = 0,
        ywTypeFrameset = 2,
        ywTypeTemplate = 1

    }
    export const enum YwDocumentType {
        ywTypeDocument = 0,
        ywTypeFrameset = 2,
        ywTypeTemplate = 1

    }
    export const enum YwDropPosition {
        ywDropNone = 0,
        ywDropNormal = 1,
        ywDropMargin = 2

    }
    export const enum YwDropPosition {
        ywDropNone = 0,
        ywDropNormal = 1,
        ywDropMargin = 2

    }
    export const enum YwEmailHTMLFidelity {
        ywEmailHTMLFidelityLow = 1,
        ywEmailHTMLFidelityMedium = 2,
        ywEmailHTMLFidelityHigh = 3

    }
    export const enum YwEmailHTMLFidelity {
        ywEmailHTMLFidelityLow = 1,
        ywEmailHTMLFidelityMedium = 2,
        ywEmailHTMLFidelityHigh = 3

    }
    export const enum YwEmphasisMark {
        ywEmphasisMarkNone = 0,
        ywEmphasisMarkOverComma = 2,
        ywEmphasisMarkOverSolidCircle = 1,
        ywEmphasisMarkOverWhiteCircle = 3,
        ywEmphasisMarkUnderSolidCircle = 4

    }
    export const enum YwEmphasisMark {
        ywEmphasisMarkNone = 0,
        ywEmphasisMarkOverComma = 2,
        ywEmphasisMarkOverSolidCircle = 1,
        ywEmphasisMarkOverWhiteCircle = 3,
        ywEmphasisMarkUnderSolidCircle = 4

    }
    export const enum YwEndnoteLocation {
        ywEndOfSection = 0,
        ywEndOfDocument = 1

    }
    export const enum YwEndnoteLocation {
        ywEndOfSection = 0,
        ywEndOfDocument = 1

    }
    export const enum YwEnvelopeOrientation {
        ywCenterClockwise = 7,
        ywCenterLandscape = 4,
        ywCenterPortrait = 1,
        ywLeftClockwise = 6,
        ywLeftLandscape = 3,
        ywLeftPortrait = 0,
        ywRightClockwise = 8,
        ywRightLandscape = 5,
        ywRightPortrait = 2

    }
    export const enum YwEnvelopeOrientation {
        ywCenterClockwise = 7,
        ywCenterLandscape = 4,
        ywCenterPortrait = 1,
        ywLeftClockwise = 6,
        ywLeftLandscape = 3,
        ywLeftPortrait = 0,
        ywRightClockwise = 8,
        ywRightLandscape = 5,
        ywRightPortrait = 2

    }
    export const enum YwFieldKind {
        ywFieldKindCold = 3,
        ywFieldKindHot = 1,
        ywFieldKindNone = 0,
        ywFieldKindWarm = 2

    }
    export const enum YwFieldKind {
        ywFieldKindCold = 3,
        ywFieldKindHot = 1,
        ywFieldKindNone = 0,
        ywFieldKindWarm = 2

    }
    export const enum YwFieldShading {
        wdFieldShadingNever = 0,
        wdFieldShadingAlways = 1,
        wdFieldShadingWhenSelected = 2

    }
    export const enum YwFieldShading {
        wdFieldShadingNever = 0,
        wdFieldShadingAlways = 1,
        wdFieldShadingWhenSelected = 2

    }
    export const enum YwFieldType {
        ywFieldAddin = 81,
        ywFieldAddressBlock = 93,
        ywFieldAdvance = 84,
        ywFieldAsk = 38,
        ywFieldAuthor = 17,
        ywFieldAutoNum = 54,
        ywFieldAutoNumLegal = 53,
        ywFieldAutoNumOutline = 52,
        ywFieldAutoText = 79,
        ywFieldAutoTextList = 89,
        ywFieldBarCode = 63,
        ywFieldBidiOutline = 92,
        ywFieldComments = 19,
        ywFieldCompare = 80,
        ywFieldCreateDate = 21,
        ywFieldData = 40,
        ywFieldDatabase = 78,
        ywFieldDate = 31,
        ywFieldDDE = 45,
        ywFieldDDEAuto = 46,
        ywFieldDocProperty = 85,
        ywFieldDocVariable = 64,
        ywFieldEditTime = 25,
        ywFieldEmbed = 58,
        ywFieldEmpty = -1,
        ywFieldExpression = 34,
        ywFieldFileName = 29,
        ywFieldFileSize = 69,
        ywFieldFillIn = 39,
        ywFieldFootnoteRef = 5,
        ywFieldFormCheckBox = 71,
        ywFieldFormDropDown = 83,
        ywFieldFormTextInput = 70,
        ywFieldFormula = 49,
        ywFieldGlossary = 47,
        ywFieldGoToButton = 50,
        ywFieldGreetingLine = 94,
        ywFieldHTMLActiveX = 91,
        ywFieldHyperlink = 88,
        ywFieldIf = 7,
        ywFieldImport = 55,
        ywFieldInclude = 36,
        ywFieldIncludePicture = 67,
        ywFieldIncludeText = 68,
        ywFieldIndex = 8,
        ywFieldIndexEntry = 4,
        ywFieldInfo = 14,
        ywFieldKeyWord = 18,
        ywFieldLastSavedBy = 20,
        ywFieldLink = 56,
        ywFieldListNum = 90,
        ywFieldMacroButton = 51,
        ywFieldMergeField = 59,
        ywFieldMergeRec = 44,
        ywFieldMergeSeq = 75,
        ywFieldNext = 41,
        ywFieldNextIf = 42,
        ywFieldNoteRef = 72,
        ywFieldNumChars = 28,
        ywFieldNumPages = 26,
        ywFieldNumWords = 27,
        ywFieldOCX = 87,
        ywFieldPage = 33,
        ywFieldPageRef = 37,
        ywFieldPrint = 48,
        ywFieldPrintDate = 23,
        ywFieldPrivate = 77,
        ywFieldQuote = 35,
        ywFieldRef = 3,
        ywFieldRefDoc = 11,
        ywFieldRevisionNum = 24,
        ywFieldSaveDate = 22,
        ywFieldSection = 65,
        ywFieldSectionPages = 66,
        ywFieldSequence = 12,
        ywFieldSet = 6,
        ywFieldShape = 95,
        ywFieldSkipIf = 43,
        ywFieldStyleRef = 10,
        ywFieldSubject = 16,
        ywFieldSubscriber = 82,
        ywFieldSymbol = 57,
        ywFieldTemplate = 30,
        ywFieldTime = 32,
        ywFieldTitle = 15,
        ywFieldTOA = 73,
        ywFieldTOAEntry = 74,
        ywFieldTOC = 13,
        ywFieldTOCEntry = 9,
        ywFieldUserAddress = 62,
        ywFieldUserInitials = 61,
        ywFieldUserName = 60

    }
    export const enum YwFieldType {
        ywFieldAddin = 81,
        ywFieldAddressBlock = 93,
        ywFieldAdvance = 84,
        ywFieldAsk = 38,
        ywFieldAuthor = 17,
        ywFieldAutoNum = 54,
        ywFieldAutoNumLegal = 53,
        ywFieldAutoNumOutline = 52,
        ywFieldAutoText = 79,
        ywFieldAutoTextList = 89,
        ywFieldBarCode = 63,
        ywFieldBidiOutline = 92,
        ywFieldComments = 19,
        ywFieldCompare = 80,
        ywFieldCreateDate = 21,
        ywFieldData = 40,
        ywFieldDatabase = 78,
        ywFieldDate = 31,
        ywFieldDDE = 45,
        ywFieldDDEAuto = 46,
        ywFieldDocProperty = 85,
        ywFieldDocVariable = 64,
        ywFieldEditTime = 25,
        ywFieldEmbed = 58,
        ywFieldEmpty = -1,
        ywFieldExpression = 34,
        ywFieldFileName = 29,
        ywFieldFileSize = 69,
        ywFieldFillIn = 39,
        ywFieldFootnoteRef = 5,
        ywFieldFormCheckBox = 71,
        ywFieldFormDropDown = 83,
        ywFieldFormTextInput = 70,
        ywFieldFormula = 49,
        ywFieldGlossary = 47,
        ywFieldGoToButton = 50,
        ywFieldGreetingLine = 94,
        ywFieldHTMLActiveX = 91,
        ywFieldHyperlink = 88,
        ywFieldIf = 7,
        ywFieldImport = 55,
        ywFieldInclude = 36,
        ywFieldIncludePicture = 67,
        ywFieldIncludeText = 68,
        ywFieldIndex = 8,
        ywFieldIndexEntry = 4,
        ywFieldInfo = 14,
        ywFieldKeyWord = 18,
        ywFieldLastSavedBy = 20,
        ywFieldLink = 56,
        ywFieldListNum = 90,
        ywFieldMacroButton = 51,
        ywFieldMergeField = 59,
        ywFieldMergeRec = 44,
        ywFieldMergeSeq = 75,
        ywFieldNext = 41,
        ywFieldNextIf = 42,
        ywFieldNoteRef = 72,
        ywFieldNumChars = 28,
        ywFieldNumPages = 26,
        ywFieldNumWords = 27,
        ywFieldOCX = 87,
        ywFieldPage = 33,
        ywFieldPageRef = 37,
        ywFieldPrint = 48,
        ywFieldPrintDate = 23,
        ywFieldPrivate = 77,
        ywFieldQuote = 35,
        ywFieldRef = 3,
        ywFieldRefDoc = 11,
        ywFieldRevisionNum = 24,
        ywFieldSaveDate = 22,
        ywFieldSection = 65,
        ywFieldSectionPages = 66,
        ywFieldSequence = 12,
        ywFieldSet = 6,
        ywFieldShape = 95,
        ywFieldSkipIf = 43,
        ywFieldStyleRef = 10,
        ywFieldSubject = 16,
        ywFieldSubscriber = 82,
        ywFieldSymbol = 57,
        ywFieldTemplate = 30,
        ywFieldTime = 32,
        ywFieldTitle = 15,
        ywFieldTOA = 73,
        ywFieldTOAEntry = 74,
        ywFieldTOC = 13,
        ywFieldTOCEntry = 9,
        ywFieldUserAddress = 62,
        ywFieldUserInitials = 61,
        ywFieldUserName = 60

    }
    export const enum YwFindWrap {
        ywFindAsk = 2,
        ywFindContinue = 1,
        ywFindStop = 0

    }
    export const enum YwFindWrap {
        ywFindAsk = 2,
        ywFindContinue = 1,
        ywFindStop = 0

    }
    export const enum YwFootnoteLocation {
        ywBeneathText = 1,
        ywBottomOfPage = 0

    }
    export const enum YwFootnoteLocation {
        ywBeneathText = 1,
        ywBottomOfPage = 0

    }
    export const enum YwFramesetSizeType {
        ywFramesetSizeTypePercent = 0,
        ywFramesetSizeTypeFixed = 1,
        ywFramesetSizeTypeRelative = 2

    }
    export const enum YwFramesetSizeType {
        ywFramesetSizeTypePercent = 0,
        ywFramesetSizeTypeFixed = 1,
        ywFramesetSizeTypeRelative = 2

    }
    export const enum YwFrameSizeRule {
        ywFrameAtLeast = 1,
        ywFrameExact = 2,
        ywFrameAuto = 0

    }
    export const enum YwFrameSizeRule {
        ywFrameAtLeast = 1,
        ywFrameExact = 2,
        ywFrameAuto = 0

    }
    export const enum YwGoToDirection {
        ywGoToAbsolute = 1,
        ywGoToFirst = 1,
        ywGoToLast = -1,
        ywGoToNext = 2,
        ywGoToPrevious = 3,
        ywGoToRelative = 2

    }
    export const enum YwGoToDirection {
        ywGoToAbsolute = 1,
        ywGoToFirst = 1,
        ywGoToLast = -1,
        ywGoToNext = 2,
        ywGoToPrevious = 3,
        ywGoToRelative = 2

    }
    export const enum YwGoToItem {
        ywGoToBookmark = -1,
        ywGoToComment = 6,
        ywGoToEndnote = 5,
        ywGoToEquation = 10,
        ywGoToField = 7,
        ywGoToFootnote = 4,
        ywGoToGrammaticalError = 14,
        ywGoToGraphic = 8,
        ywGoToHeading = 11,
        ywGoToLine = 3,
        ywGoToObject = 9,
        ywGoToPage = 1,
        ywGoToPercent = 12,
        ywGoToProofreadingError = 15,
        ywGoToSection = 0,
        ywGoToSpellingError = 13,
        ywGoToTable = 2

    }
    export const enum YwGoToItem {
        ywGoToBookmark = -1,
        ywGoToComment = 6,
        ywGoToEndnote = 5,
        ywGoToEquation = 10,
        ywGoToField = 7,
        ywGoToFootnote = 4,
        ywGoToGrammaticalError = 14,
        ywGoToGraphic = 8,
        ywGoToHeading = 11,
        ywGoToLine = 3,
        ywGoToObject = 9,
        ywGoToPage = 1,
        ywGoToPercent = 12,
        ywGoToProofreadingError = 15,
        ywGoToSection = 0,
        ywGoToSpellingError = 13,
        ywGoToTable = 2

    }
    export const enum YwGranularity {

    }
    export const enum YwGranularity {

    }
    export const enum YwGutterStyle {
        ywGutterPosLeft = 0,
        ywGutterPosTop = 1,
        ywGutterPosRight = 2

    }
    export const enum YwGutterStyle {
        ywGutterPosLeft = 0,
        ywGutterPosTop = 1,
        ywGutterPosRight = 2

    }
    export const enum YwGutterStyleOld {
        ywGutterStyleBidi = 2,
        ywGutterStyleLatin = -10

    }
    export const enum YwGutterStyleOld {
        ywGutterStyleBidi = 2,
        ywGutterStyleLatin = -10

    }
    export const enum YwHeaderFooterIndex {
        ywHeaderFooterEvenPages = 3,
        ywHeaderFooterFirstPage = 2,
        ywHeaderFooterPrimary = 1

    }
    export const enum YwHeaderFooterIndex {
        ywHeaderFooterEvenPages = 3,
        ywHeaderFooterFirstPage = 2,
        ywHeaderFooterPrimary = 1

    }
    export const enum YwHighAnsiText {
        ywAutoDetectHighAnsiFarEast = 2,
        ywHighAnsiIsHighAnsi = 1,
        ywHighAnsiIsFarEast = 0

    }
    export const enum YwHighAnsiText {
        ywAutoDetectHighAnsiFarEast = 2,
        ywHighAnsiIsHighAnsi = 1,
        ywHighAnsiIsFarEast = 0

    }
    export const enum YwHorizontalLineAlignment {
        ywHorizontalLineAlignCenter = 1,
        ywHorizontalLineAlignRight = 2,
        ywHorizontalLineAlignLeft = 0

    }
    export const enum YwHorizontalLineAlignment {
        ywHorizontalLineAlignCenter = 1,
        ywHorizontalLineAlignRight = 2,
        ywHorizontalLineAlignLeft = 0

    }
    export const enum YwInformation {
        ywActiveEndAdjustedPageNumber = 1,
        ywActiveEndSectionNumber = 2,
        ywActiveEndPageNumber = 3,
        ywNumberOfPagesInDocument = 4,
        ywHorizontalPositionRelativeToPage = 5,
        ywVerticalPositionRelativeToPage = 6,
        ywHorizontalPositionRelativeToTextBoundary = 7,
        ywVerticalPositionRelativeToTextBoundary = 8,
        ywFirstCharacterColumnNumber = 9,
        ywFirstCharacterLineNumber = 10,
        ywFrameIsSelected = 11,
        ywWithInTable = 12,
        ywStartOfRangeRowNumber = 13,
        ywEndOfRangeRowNumber = 14,
        ywMaximumNumberOfRows = 15,
        ywStartOfRangeColumnNumber = 16,
        ywEndOfRangeColumnNumber = 17,
        ywMaximumNumberOfColumns = 18,
        ywZoomPercentage = 19,
        ywSelectionMode = 20,
        ywCapsLock = 21,
        ywNumLock = 22,
        ywOverType = 23,
        ywRevisionMarking = 24,
        ywInFootnoteEndnotePane = 25,
        ywInCommentPane = 26,
        ywInHeaderFooter = 28,
        ywAtEndOfRowMarker = 31,
        ywReferenceOfType = 32,
        ywHeaderFooterType = 33,
        ywInMasterDocument = 34,
        ywInFootnote = 35,
        ywInEndnote = 36,
        ywInWordMail = 37,
        ywInClipboard = 38

    }
    export const enum YwInformation {
        ywActiveEndAdjustedPageNumber = 1,
        ywActiveEndSectionNumber = 2,
        ywActiveEndPageNumber = 3,
        ywNumberOfPagesInDocument = 4,
        ywHorizontalPositionRelativeToPage = 5,
        ywVerticalPositionRelativeToPage = 6,
        ywHorizontalPositionRelativeToTextBoundary = 7,
        ywVerticalPositionRelativeToTextBoundary = 8,
        ywFirstCharacterColumnNumber = 9,
        ywFirstCharacterLineNumber = 10,
        ywFrameIsSelected = 11,
        ywWithInTable = 12,
        ywStartOfRangeRowNumber = 13,
        ywEndOfRangeRowNumber = 14,
        ywMaximumNumberOfRows = 15,
        ywStartOfRangeColumnNumber = 16,
        ywEndOfRangeColumnNumber = 17,
        ywMaximumNumberOfColumns = 18,
        ywZoomPercentage = 19,
        ywSelectionMode = 20,
        ywCapsLock = 21,
        ywNumLock = 22,
        ywOverType = 23,
        ywRevisionMarking = 24,
        ywInFootnoteEndnotePane = 25,
        ywInCommentPane = 26,
        ywInHeaderFooter = 28,
        ywAtEndOfRowMarker = 31,
        ywReferenceOfType = 32,
        ywHeaderFooterType = 33,
        ywInMasterDocument = 34,
        ywInFootnote = 35,
        ywInEndnote = 36,
        ywInWordMail = 37,
        ywInClipboard = 38

    }
    export const enum YwInlineShapeType {
        ywInlineShapeChart = 12,
        ywInlineShapeDiagram = 13,
        ywInlineShapeEmbeddedOLEObject = 1,
        ywInlineShapeHorizontalLine = 6,
        ywInlineShapeLinkedOLEObject = 2,
        ywInlineShapeLinkedPicture = 4,
        ywInlineShapeLinkedPictureHorizontalLine = 8,
        ywInlineShapeLockedCanvas = 14,
        ywInlineShapeOLEControlObject = 5,
        ywInlineShapeOWSAnchor = 11,
        ywInlineShapePicture = 3,
        ywInlineShapePictureBullet = 9,
        ywInlineShapePictureHorizontalLine = 7,
        ywInlineShapeScriptAnchor = 10

    }
    export const enum YwInlineShapeType {
        ywInlineShapeChart = 12,
        ywInlineShapeDiagram = 13,
        ywInlineShapeEmbeddedOLEObject = 1,
        ywInlineShapeHorizontalLine = 6,
        ywInlineShapeLinkedOLEObject = 2,
        ywInlineShapeLinkedPicture = 4,
        ywInlineShapeLinkedPictureHorizontalLine = 8,
        ywInlineShapeLockedCanvas = 14,
        ywInlineShapeOLEControlObject = 5,
        ywInlineShapeOWSAnchor = 11,
        ywInlineShapePicture = 3,
        ywInlineShapePictureBullet = 9,
        ywInlineShapePictureHorizontalLine = 7,
        ywInlineShapeScriptAnchor = 10

    }
    export const enum YwInsertCells {
        ywInsertCellsEntireColumn = 3,
        ywInsertCellsEntireRow = 2,
        ywInsertCellsShiftDown = 1,
        ywInsertCellsShiftRight = 0

    }
    export const enum YwInsertCells {
        ywInsertCellsEntireColumn = 3,
        ywInsertCellsEntireRow = 2,
        ywInsertCellsShiftDown = 1,
        ywInsertCellsShiftRight = 0

    }
    export const enum YwInsertedTextMark {
        ywInsertedTextMarkNone = 0,
        ywInsertedTextMarkBold = 1,
        ywInsertedTextMarkItalic = 2,
        ywInsertedTextMarkUnderline = 3,
        ywInsertedTextMarkDoubleUnderline = 4,
        ywInsertedTextMarkColorOnly = 5,
        ywInsertedTextMarkStrikeThrough = 6

    }
    export const enum YwInsertedTextMark {
        ywInsertedTextMarkNone = 0,
        ywInsertedTextMarkBold = 1,
        ywInsertedTextMarkItalic = 2,
        ywInsertedTextMarkUnderline = 3,
        ywInsertedTextMarkDoubleUnderline = 4,
        ywInsertedTextMarkColorOnly = 5,
        ywInsertedTextMarkStrikeThrough = 6

    }
    export const enum YwJustificationMode {
        ywJustificationModeCompress = 1,
        ywJustificationModeCompressKana = 2,
        ywJustificationModeExpand = 0

    }
    export const enum YwJustificationMode {
        ywJustificationModeCompress = 1,
        ywJustificationModeCompressKana = 2,
        ywJustificationModeExpand = 0

    }
    export const enum YwKey {
        ywKey0 = 48,
        ywKey1 = 49,
        ywKey2 = 50,
        ywKey3 = 51,
        ywKey4 = 52,
        ywKey5 = 53,
        ywKey6 = 54,
        ywKey7 = 55,
        ywKey8 = 56,
        ywKey9 = 57,
        ywKeyA = 65,
        ywKeyB = 66,
        ywKeyC = 67,
        ywKeyD = 68,
        ywKeyE = 69,
        ywKeyF = 70,
        ywKeyG = 71,
        ywKeyH = 72,
        ywKeyI = 73,
        ywKeyJ = 74,
        ywKeyK = 75,
        ywKeyL = 76,
        ywKeyM = 77,
        ywKeyN = 78,
        ywKeyO = 79,
        ywKeyP = 80,
        ywKeyQ = 81,
        ywKeyR = 82,
        ywKeyS = 83,
        ywKeyT = 84,
        ywKeyU = 85,
        ywKeyV = 86,
        ywKeyW = 87,
        ywKeyX = 88,
        ywKeyY = 89,
        ywKeyZ = 90,
        ywKeyF1 = 112,
        ywKeyF2 = 113,
        ywKeyF3 = 114,
        ywKeyF4 = 115,
        ywKeyF5 = 116,
        ywKeyF6 = 117,
        ywKeyF7 = 118,
        ywKeyF8 = 119,
        ywKeyF9 = 120,
        ywKeyF10 = 121,
        ywKeyF11 = 122,
        ywKeyF12 = 123,
        ywKeyF13 = 124,
        ywKeyF14 = 125,
        ywKeyF15 = 126,
        ywKeyF16 = 127,
        ywKeyNumeric0 = 96,
        ywKeyNumeric1 = 97,
        ywKeyNumeric2 = 98,
        ywKeyNumeric3 = 99,
        ywKeyNumeric4 = 100,
        ywKeyNumeric5 = 101,
        ywKeyNumeric5Special = 12,
        ywKeyNumeric6 = 102,
        ywKeyNumeric7 = 103,
        ywKeyNumeric8 = 104,
        ywKeyNumeric9 = 105,
        ywKeyNumericMultiply = 106,
        ywKeyNumericAdd = 107,
        ywKeyNumericSubtract = 109,
        ywKeyNumericDecimal = 110,
        ywKeyNumericDivide = 111,
        ywKeyAlt = 1024,
        ywKeyBackSingleQuote = 192,
        ywKeyBackSlash = 220,
        ywKeyBackspace = 8,
        ywKeyCloseSquareBrace = 221,
        ywKeyComma = 188,
        ywKeyCommand = 512,
        ywKeyControl = 512,
        ywKeyDelete = 46,
        ywKeyEnd = 35,
        ywKeyEquals = 187,
        ywKeyEsc = 27,
        ywKeyHome = 36,
        ywKeyHyphen = 189,
        ywKeyInsert = 45,
        ywKeyOpenSquareBrace = 219,
        ywKeyOption = 1024,
        ywKeyPageDown = 34,
        ywKeyPageUp = 33,
        ywKeyPause = 19,
        ywKeyPeriod = 190,
        ywKeyReturn = 13,
        ywKeyScrollLock = 145,
        ywKeySemiColon = 186,
        ywKeyShift = 256,
        ywKeySingleQuote = 222,
        ywKeySlash = 191,
        ywKeySpacebar = 32,
        ywKeyTab = 9,
        ywNokey = 255

    }
    export const enum YwKey {
        ywKey0 = 48,
        ywKey1 = 49,
        ywKey2 = 50,
        ywKey3 = 51,
        ywKey4 = 52,
        ywKey5 = 53,
        ywKey6 = 54,
        ywKey7 = 55,
        ywKey8 = 56,
        ywKey9 = 57,
        ywKeyA = 65,
        ywKeyB = 66,
        ywKeyC = 67,
        ywKeyD = 68,
        ywKeyE = 69,
        ywKeyF = 70,
        ywKeyG = 71,
        ywKeyH = 72,
        ywKeyI = 73,
        ywKeyJ = 74,
        ywKeyK = 75,
        ywKeyL = 76,
        ywKeyM = 77,
        ywKeyN = 78,
        ywKeyO = 79,
        ywKeyP = 80,
        ywKeyQ = 81,
        ywKeyR = 82,
        ywKeyS = 83,
        ywKeyT = 84,
        ywKeyU = 85,
        ywKeyV = 86,
        ywKeyW = 87,
        ywKeyX = 88,
        ywKeyY = 89,
        ywKeyZ = 90,
        ywKeyF1 = 112,
        ywKeyF2 = 113,
        ywKeyF3 = 114,
        ywKeyF4 = 115,
        ywKeyF5 = 116,
        ywKeyF6 = 117,
        ywKeyF7 = 118,
        ywKeyF8 = 119,
        ywKeyF9 = 120,
        ywKeyF10 = 121,
        ywKeyF11 = 122,
        ywKeyF12 = 123,
        ywKeyF13 = 124,
        ywKeyF14 = 125,
        ywKeyF15 = 126,
        ywKeyF16 = 127,
        ywKeyNumeric0 = 96,
        ywKeyNumeric1 = 97,
        ywKeyNumeric2 = 98,
        ywKeyNumeric3 = 99,
        ywKeyNumeric4 = 100,
        ywKeyNumeric5 = 101,
        ywKeyNumeric5Special = 12,
        ywKeyNumeric6 = 102,
        ywKeyNumeric7 = 103,
        ywKeyNumeric8 = 104,
        ywKeyNumeric9 = 105,
        ywKeyNumericMultiply = 106,
        ywKeyNumericAdd = 107,
        ywKeyNumericSubtract = 109,
        ywKeyNumericDecimal = 110,
        ywKeyNumericDivide = 111,
        ywKeyAlt = 1024,
        ywKeyBackSingleQuote = 192,
        ywKeyBackSlash = 220,
        ywKeyBackspace = 8,
        ywKeyCloseSquareBrace = 221,
        ywKeyComma = 188,
        ywKeyCommand = 512,
        ywKeyControl = 512,
        ywKeyDelete = 46,
        ywKeyEnd = 35,
        ywKeyEquals = 187,
        ywKeyEsc = 27,
        ywKeyHome = 36,
        ywKeyHyphen = 189,
        ywKeyInsert = 45,
        ywKeyOpenSquareBrace = 219,
        ywKeyOption = 1024,
        ywKeyPageDown = 34,
        ywKeyPageUp = 33,
        ywKeyPause = 19,
        ywKeyPeriod = 190,
        ywKeyReturn = 13,
        ywKeyScrollLock = 145,
        ywKeySemiColon = 186,
        ywKeyShift = 256,
        ywKeySingleQuote = 222,
        ywKeySlash = 191,
        ywKeySpacebar = 32,
        ywKeyTab = 9,
        ywNokey = 255

    }
    export const enum YwKeyCategory {
        ywKeyCategoryAutoText = 4,
        ywKeyCategoryCommand = 1,
        ywKeyCategoryDisable = 0,
        ywKeyCategoryFont = 3,
        ywKeyCategoryMacro = 2,
        ywKeyCategoryNil = -1,
        ywKeyCategoryPrefix = 7,
        ywKeyCategoryStyle = 5,
        ywKeyCategorySymbol = 6

    }
    export const enum YwKeyCategory {
        ywKeyCategoryAutoText = 4,
        ywKeyCategoryCommand = 1,
        ywKeyCategoryDisable = 0,
        ywKeyCategoryFont = 3,
        ywKeyCategoryMacro = 2,
        ywKeyCategoryNil = -1,
        ywKeyCategoryPrefix = 7,
        ywKeyCategoryStyle = 5,
        ywKeyCategorySymbol = 6

    }
    export const enum YwLanguageID {
        ywAfrikaans = 1078,
        ywAlbanian = 1052,
        ywAmharic = 1118,
        ywArabic = 1025,
        ywArabicAlgeria = 5121,
        ywArabicBahrain = 15361,
        ywArabicEgypt = 3073,
        ywArabicIraq = 2049,
        ywArabicJordan = 11265,
        ywArabicKuwait = 13313,
        ywArabicLebanon = 12289,
        ywArabicLibya = 4097,
        ywArabicMorocco = 6145,
        ywArabicOman = 8193,
        ywArabicQatar = 16385,
        ywArabicSyria = 10241,
        ywArabicTunisia = 7169,
        ywArabicUAE = 14337,
        ywArabicYemen = 9217,
        ywArmenian = 1067,
        ywAssamese = 1101,
        ywAzeriCyrillic = 2092,
        ywAzeriLatin = 1068,
        ywBasque = 1069,
        ywBelgianDutch = 2067,
        ywBelgianFrench = 2060,
        ywBengali = 1093,
        ywBrazilianPortuguese = 1046,
        ywBulgarian = 1026,
        ywBurmese = 1109,
        ywByelorussian = 1059,
        ywCatalan = 1027,
        ywCherokee = 1116,
        ywChineseHongKongSAR = 3076,
        ywChineseMacaoSAR = 5124,
        ywChineseSingapore = 4100,
        ywCroatian = 1050,
        ywCzech = 1029,
        ywDanish = 1030,
        ywDivehi = 1125,
        ywDutch = 1043,
        ywDzongkhaBhutan = 2129,
        ywEdo = 1126,
        ywEnglishAUS = 3081,
        ywEnglishBelize = 10249,
        ywEnglishCanadian = 4105,
        ywEnglishCaribbean = 9225,
        ywEnglishIndonesia = 14345,
        ywEnglishIreland = 6153,
        ywEnglishJamaica = 8201,
        ywEnglishNewZealand = 5129,
        ywEnglishPhilippines = 13321,
        ywEnglishSouthAfrica = 7177,
        ywEnglishTrinidadTobago = 11273,
        ywEnglishUK = 2057,
        ywEnglishUS = 1033,
        ywEnglishZimbabwe = 12297,
        ywEstonian = 1061,
        ywFaeroese = 1080,
        ywFarsi = 1065,
        ywFilipino = 1124,
        ywFinnish = 1035,
        ywFrench = 1036,
        ywFrenchCameroon = 11276,
        ywFrenchCanadian = 3084,
        ywFrenchCotedIvoire = 12300,
        ywFrenchHaiti = 15372,
        ywFrenchLuxembourg = 5132,
        ywFrenchMali = 13324,
        ywFrenchMonaco = 6156,
        ywFrenchMorocco = 14348,
        ywFrenchReunion = 8204,
        ywFrenchSenegal = 10252,
        ywFrenchWestIndies = 7180,
        ywFrenchZaire = 9228,
        ywFrisianNetherlands = 1122,
        ywFulfulde = 1127,
        ywGaelicIreland = 2108,
        ywGaelicScotland = 1084,
        ywGalician = 1110,
        ywGeorgian = 1079,
        ywGerman = 1031,
        ywGermanAustria = 3079,
        ywGermanLiechtenstein = 5127,
        ywGermanLuxembourg = 4103,
        ywGreek = 1032,
        ywGuarani = 1140,
        ywGujarati = 1095,
        ywHausa = 1128,
        ywHawaiian = 1141,
        ywHebrew = 1037,
        ywHindi = 1081,
        ywHungarian = 1038,
        ywIbibio = 1129,
        ywIcelandic = 1039,
        ywIgbo = 1136,
        ywIndonesian = 1057,
        ywInuktitut = 1117,
        ywItalian = 1040,
        ywJapanese = 1041,
        ywKannada = 1099,
        ywKanuri = 1137,
        ywKashmiri = 1120,
        ywKazakh = 1087,
        ywKhmer = 1107,
        ywKirghiz = 1088,
        ywKonkani = 1111,
        ywKorean = 1042,
        ywKyrgyz = 1088,
        ywLanguageNone = 0,
        ywLao = 1108,
        ywLatin = 1142,
        ywLatvian = 1062,
        ywLithuanian = 1063,
        ywMacedonian = 1071,
        ywMalayalam = 1100,
        ywMalayBruneiDarussalam = 2110,
        ywMalaysian = 1086,
        ywMaltese = 1082,
        ywManipuri = 1112,
        ywMarathi = 1102,
        ywMexicanSpanish = 2058,
        ywMongolian = 1104,
        ywNepali = 1121,
        ywNoProofing = 1024,
        ywNorwegianBokmol = 1044,
        ywNorwegianNynorsk = 2068,
        ywOriya = 1096,
        ywOromo = 1138,
        ywPashto = 1123,
        ywPolish = 1045,
        ywPortuguese = 2070,
        ywPunjabi = 1094,
        ywRhaetoRomanic = 1047,
        ywRomanian = 1048,
        ywRomanianMoldova = 2072,
        ywRussian = 1049,
        ywRussianMoldova = 2073,
        ywSamiLappish = 1083,
        ywSanskrit = 1103,
        ywSerbianCyrillic = 3098,
        ywSerbianLatin = 2074,
        ywSesotho = 1072,
        ywSimplifiedChinese = 2052,
        ywSindhi = 1113,
        ywSindhiPakistan = 2137,
        ywSinhalese = 1115,
        ywSlovak = 1051,
        ywSlovenian = 1060,
        ywSomali = 1143,
        ywSorbian = 1070,
        ywSpanish = 1034,
        ywSpanishArgentina = 11274,
        ywSpanishBolivia = 16394,
        ywSpanishChile = 13322,
        ywSpanishColombia = 9226,
        ywSpanishCostaRica = 5130,
        ywSpanishDominicanRepublic = 7178,
        ywSpanishEcuador = 12298,
        ywSpanishElSalvador = 17418,
        ywSpanishGuatemala = 4106,
        ywSpanishHonduras = 18442,
        ywSpanishModernSort = 3082,
        ywSpanishNicaragua = 19466,
        ywSpanishPanama = 6154,
        ywSpanishParaguay = 15370,
        ywSpanishPeru = 10250,
        ywSpanishPuertoRico = 20490,
        ywSpanishUruguay = 14346,
        ywSpanishVenezuela = 8202,
        ywSutu = 1072,
        ywSwahili = 1089,
        ywSwedish = 1053,
        ywSwedishFinland = 2077,
        ywSwissFrench = 4108,
        ywSwissGerman = 2055,
        ywSwissItalian = 2064,
        ywSyriac = 1114,
        ywTajik = 1064,
        ywTamazight = 1119,
        ywTamazightLatin = 2143,
        ywTamil = 1097,
        ywTatar = 1092,
        ywTelugu = 1098,
        ywThai = 1054,
        ywTibetan = 1105,
        ywTigrignaEritrea = 2163,
        ywTigrignaEthiopic = 1139,
        ywTraditionalChinese = 1028,
        ywTsonga = 1073,
        ywTswana = 1074,
        ywTurkish = 1055,
        ywTurkmen = 1090,
        ywUkrainian = 1058,
        ywUrdu = 1056,
        ywUzbekCyrillic = 2115,
        ywUzbekLatin = 1091,
        ywVenda = 1075,
        ywVietnamese = 1066,
        ywWelsh = 1106,
        ywXhosa = 1076,
        ywYi = 1144,
        ywYiddish = 1085,
        ywYoruba = 1130,
        ywZulu = 1077

    }
    export const enum YwLanguageID {
        ywAfrikaans = 1078,
        ywAlbanian = 1052,
        ywAmharic = 1118,
        ywArabic = 1025,
        ywArabicAlgeria = 5121,
        ywArabicBahrain = 15361,
        ywArabicEgypt = 3073,
        ywArabicIraq = 2049,
        ywArabicJordan = 11265,
        ywArabicKuwait = 13313,
        ywArabicLebanon = 12289,
        ywArabicLibya = 4097,
        ywArabicMorocco = 6145,
        ywArabicOman = 8193,
        ywArabicQatar = 16385,
        ywArabicSyria = 10241,
        ywArabicTunisia = 7169,
        ywArabicUAE = 14337,
        ywArabicYemen = 9217,
        ywArmenian = 1067,
        ywAssamese = 1101,
        ywAzeriCyrillic = 2092,
        ywAzeriLatin = 1068,
        ywBasque = 1069,
        ywBelgianDutch = 2067,
        ywBelgianFrench = 2060,
        ywBengali = 1093,
        ywBrazilianPortuguese = 1046,
        ywBulgarian = 1026,
        ywBurmese = 1109,
        ywByelorussian = 1059,
        ywCatalan = 1027,
        ywCherokee = 1116,
        ywChineseHongKongSAR = 3076,
        ywChineseMacaoSAR = 5124,
        ywChineseSingapore = 4100,
        ywCroatian = 1050,
        ywCzech = 1029,
        ywDanish = 1030,
        ywDivehi = 1125,
        ywDutch = 1043,
        ywDzongkhaBhutan = 2129,
        ywEdo = 1126,
        ywEnglishAUS = 3081,
        ywEnglishBelize = 10249,
        ywEnglishCanadian = 4105,
        ywEnglishCaribbean = 9225,
        ywEnglishIndonesia = 14345,
        ywEnglishIreland = 6153,
        ywEnglishJamaica = 8201,
        ywEnglishNewZealand = 5129,
        ywEnglishPhilippines = 13321,
        ywEnglishSouthAfrica = 7177,
        ywEnglishTrinidadTobago = 11273,
        ywEnglishUK = 2057,
        ywEnglishUS = 1033,
        ywEnglishZimbabwe = 12297,
        ywEstonian = 1061,
        ywFaeroese = 1080,
        ywFarsi = 1065,
        ywFilipino = 1124,
        ywFinnish = 1035,
        ywFrench = 1036,
        ywFrenchCameroon = 11276,
        ywFrenchCanadian = 3084,
        ywFrenchCotedIvoire = 12300,
        ywFrenchHaiti = 15372,
        ywFrenchLuxembourg = 5132,
        ywFrenchMali = 13324,
        ywFrenchMonaco = 6156,
        ywFrenchMorocco = 14348,
        ywFrenchReunion = 8204,
        ywFrenchSenegal = 10252,
        ywFrenchWestIndies = 7180,
        ywFrenchZaire = 9228,
        ywFrisianNetherlands = 1122,
        ywFulfulde = 1127,
        ywGaelicIreland = 2108,
        ywGaelicScotland = 1084,
        ywGalician = 1110,
        ywGeorgian = 1079,
        ywGerman = 1031,
        ywGermanAustria = 3079,
        ywGermanLiechtenstein = 5127,
        ywGermanLuxembourg = 4103,
        ywGreek = 1032,
        ywGuarani = 1140,
        ywGujarati = 1095,
        ywHausa = 1128,
        ywHawaiian = 1141,
        ywHebrew = 1037,
        ywHindi = 1081,
        ywHungarian = 1038,
        ywIbibio = 1129,
        ywIcelandic = 1039,
        ywIgbo = 1136,
        ywIndonesian = 1057,
        ywInuktitut = 1117,
        ywItalian = 1040,
        ywJapanese = 1041,
        ywKannada = 1099,
        ywKanuri = 1137,
        ywKashmiri = 1120,
        ywKazakh = 1087,
        ywKhmer = 1107,
        ywKirghiz = 1088,
        ywKonkani = 1111,
        ywKorean = 1042,
        ywKyrgyz = 1088,
        ywLanguageNone = 0,
        ywLao = 1108,
        ywLatin = 1142,
        ywLatvian = 1062,
        ywLithuanian = 1063,
        ywMacedonian = 1071,
        ywMalayalam = 1100,
        ywMalayBruneiDarussalam = 2110,
        ywMalaysian = 1086,
        ywMaltese = 1082,
        ywManipuri = 1112,
        ywMarathi = 1102,
        ywMexicanSpanish = 2058,
        ywMongolian = 1104,
        ywNepali = 1121,
        ywNoProofing = 1024,
        ywNorwegianBokmol = 1044,
        ywNorwegianNynorsk = 2068,
        ywOriya = 1096,
        ywOromo = 1138,
        ywPashto = 1123,
        ywPolish = 1045,
        ywPortuguese = 2070,
        ywPunjabi = 1094,
        ywRhaetoRomanic = 1047,
        ywRomanian = 1048,
        ywRomanianMoldova = 2072,
        ywRussian = 1049,
        ywRussianMoldova = 2073,
        ywSamiLappish = 1083,
        ywSanskrit = 1103,
        ywSerbianCyrillic = 3098,
        ywSerbianLatin = 2074,
        ywSesotho = 1072,
        ywSimplifiedChinese = 2052,
        ywSindhi = 1113,
        ywSindhiPakistan = 2137,
        ywSinhalese = 1115,
        ywSlovak = 1051,
        ywSlovenian = 1060,
        ywSomali = 1143,
        ywSorbian = 1070,
        ywSpanish = 1034,
        ywSpanishArgentina = 11274,
        ywSpanishBolivia = 16394,
        ywSpanishChile = 13322,
        ywSpanishColombia = 9226,
        ywSpanishCostaRica = 5130,
        ywSpanishDominicanRepublic = 7178,
        ywSpanishEcuador = 12298,
        ywSpanishElSalvador = 17418,
        ywSpanishGuatemala = 4106,
        ywSpanishHonduras = 18442,
        ywSpanishModernSort = 3082,
        ywSpanishNicaragua = 19466,
        ywSpanishPanama = 6154,
        ywSpanishParaguay = 15370,
        ywSpanishPeru = 10250,
        ywSpanishPuertoRico = 20490,
        ywSpanishUruguay = 14346,
        ywSpanishVenezuela = 8202,
        ywSutu = 1072,
        ywSwahili = 1089,
        ywSwedish = 1053,
        ywSwedishFinland = 2077,
        ywSwissFrench = 4108,
        ywSwissGerman = 2055,
        ywSwissItalian = 2064,
        ywSyriac = 1114,
        ywTajik = 1064,
        ywTamazight = 1119,
        ywTamazightLatin = 2143,
        ywTamil = 1097,
        ywTatar = 1092,
        ywTelugu = 1098,
        ywThai = 1054,
        ywTibetan = 1105,
        ywTigrignaEritrea = 2163,
        ywTigrignaEthiopic = 1139,
        ywTraditionalChinese = 1028,
        ywTsonga = 1073,
        ywTswana = 1074,
        ywTurkish = 1055,
        ywTurkmen = 1090,
        ywUkrainian = 1058,
        ywUrdu = 1056,
        ywUzbekCyrillic = 2115,
        ywUzbekLatin = 1091,
        ywVenda = 1075,
        ywVietnamese = 1066,
        ywWelsh = 1106,
        ywXhosa = 1076,
        ywYi = 1144,
        ywYiddish = 1085,
        ywYoruba = 1130,
        ywZulu = 1077

    }
    export const enum YwLayoutMode {
        ywLayoutModeDefault = 0,
        ywLayoutModeGrid = 1,
        ywLayoutModeLineGrid = 2,
        ywLayoutModeGenko = 3

    }
    export const enum YwLayoutMode {
        ywLayoutModeDefault = 0,
        ywLayoutModeGrid = 1,
        ywLayoutModeLineGrid = 2,
        ywLayoutModeGenko = 3

    }
    export const enum YwLetterheadLocation {
        ywLetterTop = 0,
        ywLetterBottom = 1,
        ywLetterLeft = 2,
        ywLetterRight = 3

    }
    export const enum YwLetterheadLocation {
        ywLetterTop = 0,
        ywLetterBottom = 1,
        ywLetterLeft = 2,
        ywLetterRight = 3

    }
    export const enum YwLetterStyle {
        ywFullBlock = 0,
        ywModifiedBlock = 1,
        ywSemiBlock = 2

    }
    export const enum YwLetterStyle {
        ywFullBlock = 0,
        ywModifiedBlock = 1,
        ywSemiBlock = 2

    }
    export const enum YwLineSpacing {
        ywLineSpace1pt5 = 1,
        ywLineSpaceAtLeast = 3,
        ywLineSpaceDouble = 2,
        ywLineSpaceExactly = 4,
        ywLineSpaceMultiple = 5,
        ywLineSpaceSingle = 0

    }
    export const enum YwLineSpacing {
        ywLineSpace1pt5 = 1,
        ywLineSpaceAtLeast = 3,
        ywLineSpaceDouble = 2,
        ywLineSpaceExactly = 4,
        ywLineSpaceMultiple = 5,
        ywLineSpaceSingle = 0

    }
    export const enum YwLineStyle {
        ywLineStyleDashDot = 5,
        ywLineStyleDashDotDot = 6,
        ywLineStyleDashDotStroked = 20,
        ywLineStyleDashLargeGap = 4,
        ywLineStyleDashSmallGap = 3,
        ywLineStyleDot = 2,
        ywLineStyleDouble = 7,
        ywLineStyleDoubleWavy = 19,
        ywLineStyleEmboss3D = 21,
        ywLineStyleEngrave3D = 22,
        ywLineStyleInset = 24,
        ywLineStyleNone = 0,
        ywLineStyleOutset = 23,
        ywLineStyleSingle = 1,
        ywLineStyleSingleWavy = 18,
        ywLineStyleThickThinLargeGap = 16,
        ywLineStyleThickThinMedGap = 13,
        ywLineStyleThickThinSmallGap = 10,
        ywLineStyleThinThickLargeGap = 15,
        ywLineStyleThinThickMedGap = 12,
        ywLineStyleThinThickSmallGap = 9,
        ywLineStyleThinThickThinLargeGap = 17,
        ywLineStyleThinThickThinMedGap = 14,
        ywLineStyleThinThickThinSmallGap = 11,
        ywLineStyleTriple = 8

    }
    export const enum YwLineStyle {
        ywLineStyleDashDot = 5,
        ywLineStyleDashDotDot = 6,
        ywLineStyleDashDotStroked = 20,
        ywLineStyleDashLargeGap = 4,
        ywLineStyleDashSmallGap = 3,
        ywLineStyleDot = 2,
        ywLineStyleDouble = 7,
        ywLineStyleDoubleWavy = 19,
        ywLineStyleEmboss3D = 21,
        ywLineStyleEngrave3D = 22,
        ywLineStyleInset = 24,
        ywLineStyleNone = 0,
        ywLineStyleOutset = 23,
        ywLineStyleSingle = 1,
        ywLineStyleSingleWavy = 18,
        ywLineStyleThickThinLargeGap = 16,
        ywLineStyleThickThinMedGap = 13,
        ywLineStyleThickThinSmallGap = 10,
        ywLineStyleThinThickLargeGap = 15,
        ywLineStyleThinThickMedGap = 12,
        ywLineStyleThinThickSmallGap = 9,
        ywLineStyleThinThickThinLargeGap = 17,
        ywLineStyleThinThickThinMedGap = 14,
        ywLineStyleThinThickThinSmallGap = 11,
        ywLineStyleTriple = 8

    }
    export const enum YwLineType {
        ywTableRow = 1,
        ywTextLine = 0

    }
    export const enum YwLineType {
        ywTableRow = 1,
        ywTextLine = 0

    }
    export const enum YwLineWidth {
        ywLineWidth025pt = 2,
        ywLineWidth050pt = 4,
        ywLineWidth075pt = 6,
        ywLineWidth100pt = 8,
        ywLineWidth150pt = 12,
        ywLineWidth225pt = 18,
        ywLineWidth300pt = 24,
        ywLineWidth450pt = 36,
        ywLineWidth600pt = 48

    }
    export const enum YwLineWidth {
        ywLineWidth025pt = 2,
        ywLineWidth050pt = 4,
        ywLineWidth075pt = 6,
        ywLineWidth100pt = 8,
        ywLineWidth150pt = 12,
        ywLineWidth225pt = 18,
        ywLineWidth300pt = 24,
        ywLineWidth450pt = 36,
        ywLineWidth600pt = 48

    }
    export const enum YwLinkType {
        ywLinkTypeDDEAuto = 7,
        ywLinkTypeDDE = 6,
        ywLinkTypeImport = 5,
        ywLinkTypeInclude = 4,
        ywLinkTypeReference = 3,
        ywLinkTypeText = 2,
        ywLinkTypePicture = 1,
        ywLinkTypeOLE = 0

    }
    export const enum YwLinkType {
        ywLinkTypeDDEAuto = 7,
        ywLinkTypeDDE = 6,
        ywLinkTypeImport = 5,
        ywLinkTypeInclude = 4,
        ywLinkTypeReference = 3,
        ywLinkTypeText = 2,
        ywLinkTypePicture = 1,
        ywLinkTypeOLE = 0

    }
    export const enum YwListApplyTo {
        ywListApplyToSelection = 2,
        ywListApplyToThisPointForward = 1,
        ywListApplyToWholeList = 0

    }
    export const enum YwListApplyTo {
        ywListApplyToSelection = 2,
        ywListApplyToThisPointForward = 1,
        ywListApplyToWholeList = 0

    }
    export const enum YwListGalleryType {
        ywBulletGallery = 1,
        ywNumberGallery = 2,
        ywOutlineNumberGallery = 3

    }
    export const enum YwListGalleryType {
        ywBulletGallery = 1,
        ywNumberGallery = 2,
        ywOutlineNumberGallery = 3

    }
    export const enum YwListLevelAlignment {
        ywListLevelAlignLeft = 0,
        ywListLevelAlignCenter = 1,
        ywListLevelAlignRight = 2

    }
    export const enum YwListLevelAlignment {
        ywListLevelAlignLeft = 0,
        ywListLevelAlignCenter = 1,
        ywListLevelAlignRight = 2

    }
    export const enum YwListType {
        ywListNoNumbering = 0,
        ywListListNumOnly = 1,
        ywListBullet = 2,
        ywListSimpleNumbering = 3,
        ywListOutlineNumbering = 4,
        ywListMixedNumbering = 5,
        ywListPictureBullet = 6

    }
    export const enum YwListType {
        ywListNoNumbering = 0,
        ywListListNumOnly = 1,
        ywListBullet = 2,
        ywListSimpleNumbering = 3,
        ywListOutlineNumbering = 4,
        ywListMixedNumbering = 5,
        ywListPictureBullet = 6

    }
    export const enum YwMailMergeActiveRecord {
        ywPreviousDataSourceRecord = -9,
        ywNextDataSourceRecord = -8,
        ywLastDataSourceRecord = -7,
        ywFirstDataSourceRecord = -6,
        ywLastRecord = -5,
        ywFirstRecord = -4,
        ywPreviousRecord = -3,
        ywNextRecord = -2,
        ywNoActiveRecord = -1

    }
    export const enum YwMailMergeActiveRecord {
        ywPreviousDataSourceRecord = -9,
        ywNextDataSourceRecord = -8,
        ywLastDataSourceRecord = -7,
        ywFirstDataSourceRecord = -6,
        ywLastRecord = -5,
        ywFirstRecord = -4,
        ywPreviousRecord = -3,
        ywNextRecord = -2,
        ywNoActiveRecord = -1

    }
    export const enum YwMailMergeComparison {
        ywMergeIfEqual = 0,
        ywMergeIfNotEqual = 1,
        ywMergeIfLessThan = 2,
        ywMergeIfGreaterThan = 3,
        ywMergeIfLessThanOrEqual = 4,
        ywMergeIfGreaterThanOrEqual = 5,
        ywMergeIfIsBlank = 6,
        ywMergeIfIsNotBlank = 7

    }
    export const enum YwMailMergeComparison {
        ywMergeIfEqual = 0,
        ywMergeIfNotEqual = 1,
        ywMergeIfLessThan = 2,
        ywMergeIfGreaterThan = 3,
        ywMergeIfLessThanOrEqual = 4,
        ywMergeIfGreaterThanOrEqual = 5,
        ywMergeIfIsBlank = 6,
        ywMergeIfIsNotBlank = 7

    }
    export const enum YwMailMergeDataSource {
        ywMergeInfoFromWord = 0,
        ywMergeInfoFromAccessDDE = 1,
        ywMergeInfoFromExcelDDE = 2,
        ywMergeInfoFromMSQueryDDE = 3,
        ywMergeInfoFromODBC = 4,
        ywMergeInfoFromODSO = 5,
        ywNoMergeInfo = -1

    }
    export const enum YwMailMergeDataSource {
        ywMergeInfoFromWord = 0,
        ywMergeInfoFromAccessDDE = 1,
        ywMergeInfoFromExcelDDE = 2,
        ywMergeInfoFromMSQueryDDE = 3,
        ywMergeInfoFromODBC = 4,
        ywMergeInfoFromODSO = 5,
        ywNoMergeInfo = -1

    }
    export const enum YwMailMergeDestination {
        ywSendToNewDocument = 0,
        ywSendToPrinter = 1,
        ywSendToEmail = 2,
        ywSendToFax = 3

    }
    export const enum YwMailMergeDestination {
        ywSendToNewDocument = 0,
        ywSendToPrinter = 1,
        ywSendToEmail = 2,
        ywSendToFax = 3

    }
    export const enum YwMailMergeMailFormat {
        ywMailFormatPlainText = 0,
        ywMailFormatHTML = 1

    }
    export const enum YwMailMergeMailFormat {
        ywMailFormatPlainText = 0,
        ywMailFormatHTML = 1

    }
    export const enum YwMailMergeMainDocType {
        ywFormLetters = 0,
        ywMailingLabels = 1,
        ywEnvelopes = 2,
        ywCatalog = 3,
        ywDirectory = 3,
        ywEMail = 4,
        ywFax = 5,
        ywNotAMergeDocument = -1

    }
    export const enum YwMailMergeMainDocType {
        ywFormLetters = 0,
        ywMailingLabels = 1,
        ywEnvelopes = 2,
        ywCatalog = 3,
        ywDirectory = 3,
        ywEMail = 4,
        ywFax = 5,
        ywNotAMergeDocument = -1

    }
    export const enum YwMailMergeState {
        ywNormalDocument = 0,
        ywMainDocumentOnly = 1,
        ywMainAndDataSource = 2,
        ywMainAndHeader = 3,
        ywMainAndSourceAndHeader = 4,
        ywDataSource = 5

    }
    export const enum YwMailMergeState {
        ywNormalDocument = 0,
        ywMainDocumentOnly = 1,
        ywMainAndDataSource = 2,
        ywMainAndHeader = 3,
        ywMainAndSourceAndHeader = 4,
        ywDataSource = 5

    }
    export const enum YwMappedDataFields {
        ywUniqueIdentifier = 1,
        ywCourtesyTitle = 2,
        ywFirstName = 3,
        ywMiddleName = 4,
        ywLastName = 5,
        ywSuffix = 6,
        ywNickname = 7,
        ywJobTitle = 8,
        ywCompany = 9,
        ywAddress1 = 10,
        ywAddress2 = 11,
        ywCity = 12,
        ywState = 13,
        ywPostalCode = 14,
        ywCountryRegion = 15,
        ywBusinessPhone = 16,
        ywBusinessFax = 17,
        ywHomePhone = 18,
        ywHomeFax = 19,
        ywEmailAddress = 20,
        ywWebPageURL = 21,
        ywSpouseCourtesyTitle = 22,
        ywSpouseFirstName = 23,
        ywSpouseMiddleName = 24,
        ywSpouseLastName = 25,
        ywSpouseNickname = 26,
        ywRubyFirstName = 27,
        ywRubyLastName = 28,
        ywAddress3 = 29,
        ywDepartment = 30

    }
    export const enum YwMappedDataFields {
        ywUniqueIdentifier = 1,
        ywCourtesyTitle = 2,
        ywFirstName = 3,
        ywMiddleName = 4,
        ywLastName = 5,
        ywSuffix = 6,
        ywNickname = 7,
        ywJobTitle = 8,
        ywCompany = 9,
        ywAddress1 = 10,
        ywAddress2 = 11,
        ywCity = 12,
        ywState = 13,
        ywPostalCode = 14,
        ywCountryRegion = 15,
        ywBusinessPhone = 16,
        ywBusinessFax = 17,
        ywHomePhone = 18,
        ywHomeFax = 19,
        ywEmailAddress = 20,
        ywWebPageURL = 21,
        ywSpouseCourtesyTitle = 22,
        ywSpouseFirstName = 23,
        ywSpouseMiddleName = 24,
        ywSpouseLastName = 25,
        ywSpouseNickname = 26,
        ywRubyFirstName = 27,
        ywRubyLastName = 28,
        ywAddress3 = 29,
        ywDepartment = 30

    }
    export const enum YwMeasurementUnits {
        ywInches = 0,
        ywCentimeters = 1,
        ywMillimeters = 2,
        ywPoints = 3,
        ywPicas = 4

    }
    export const enum YwMeasurementUnits {
        ywInches = 0,
        ywCentimeters = 1,
        ywMillimeters = 2,
        ywPoints = 3,
        ywPicas = 4

    }
    export const enum YwMenuID {
        ywFile = 30002,
        ywEdit = 30003,
        ywView = 30004,
        ywInsert = 30005,
        ywFormat = 30006,
        ywTool = 30007,
        ywTable = 30008,
        ywWindow = 30009,
        ywHelp = 30010,
        ywFileNew = 18,
        ywFileOpen = 23,
        ywFileClose = 106,
        ywFileSave = 3,
        ywFileSaveAs = 748,
        ywFileSaveAsHtml = 3823,
        ywFileSaveWorkspace = 846,
        ywFileFind = 5905,
        ywFilePermissionView = 7994,
        ywFileVersionHistory = 2522,
        ywFileWebPagePreview = 3655,
        ywFilePageSetupDialog = 247,
        ywFilePreviewPrint = 109,
        ywFilePrint = 4,
        ywFileProperties = 750,
        ywFileExit = 752,
        ywEditUndo = 128,
        ywEditRedo = 335,
        ywEditCut = 21,
        ywEditCopy = 19,
        ywEditPaste = 22,
        ywEditPasteSpecialDialog = 755,
        ywEditPasteAsHyperlink = 2787,
        ywEditClearMenuWord = 30021,
        ywEditClearFormat = 872,
        ywEditClearContent = 7024,
        ywEditSelectAll = 756,
        ywEditFind = 141,
        ywEditReplace = 313,
        ywEditGoto = 757,
        ywEditLinksToFiles = 759,
        ywViewNormalView = 224,
        ywViewWebLayoutView = 3621,
        ywViewPrintLayoutView = 287,
        ywViewFullScreenReadingView = 9377,
        ywViewMasterDocumentView = 4387,
        ywViewTaskPane = 5746,
        ywViewToolBars = 30045,
        ywViewRuler = 7391,
        ywViewShowParasTag = 3252,
        ywViewGridlines = 300,
        ywViewDocumentMap = 1714,
        ywViewHeaderAndFooter = 762,
        ywViewFootnotesEndnotes = 764,
        ywViewSign = 6547,
        ywViewHtmlSource = 3902,
        ywViewFullScreen = 178,
        ywViewZoom = 925,
        ywFormatFont = 253,
        ywFormatParagraph = 779,
        ywFormatBulletsAndNumbers = 783,
        ywFormatBordersShading = 781,
        ywFormatColumns = 9,
        ywFormatTabs = 780,
        ywFormatDropCap = 289,
        ywFormatTextDirection = 782,
        ywFormatChangeCase = 309,
        ywFormatFitText = 3510,
        ywFormatChineseLayout = 30463,
        ywFormatAsianLayoutPhoneticGuide = 3511,
        ywFormatAsianLayoutCharactersEnclose = 3969,
        ywFormatAsianLayoutHorizontalInVertical = 3963,
        ywFormatAsianLayoutCombineCharacters = 3512,
        ywFormatAsianLayoutTwoLinesInOne = 3964,
        ywFormatBackground = 30403,
        ywFormatTheme = 3623,
        ywFormatFrameset = 30452,
        ywFormatAutoFormat = 144,
        ywFormatStylesPane = 5757,
        ywFormatRevealFormatting = 6094,
        ywFormatObject = 2327,
        ywToolSpellingAndGrammar = 2566,
        ywToolResearchPane = 7343,
        ywToolLanguage = 30182,
        ywToolSetLanguage = 790,
        ywToolChineseTranslation = 3958,
        ywToolTranslation = 6111,
        ywToolThesaurus = 9056,
        ywToolHyphenation = 2800,
        ywToolWordCount = 792,
        ywToolAutoSummarize = 1709,
        ywToolPhonetic = 5764,
        ywToolShareWorkspace = 7710,
        ywToolTrackChanges = 2041,
        ywToolCompareAndCombine = 2338,
        ywToolProtect = 7116,
        ywToolWorkOnline = 30468,
        ywToolMeetingStart = 3727,
        ywToolMeetingArrange = 4179,
        ywToolWebDiscuss = 4177,
        ywToolLetterAndMail = 31171,
        ywToolMailMerge = 6070,
        ywToolShowMailMergeToolbar = 5911,
        ywToolEnvelopesAndLabels = 794,
        ywToolChineseEnvelopeWizard = 3972,
        ywToolEnglishEnvelopeWizard = 796,
        ywToolMacro = 30017,
        ywToolMacroDialog = 186,
        ywToolRecordMacro = 2780,
        ywToolMacroSecurity = 3627,
        ywToolVisualBasic = 1695,
        ywToolMSScriptEditor = 3631,
        ywToolComAddIns = 3754,
        ywToolDocumentTemplate = 751,
        ywToolAutoCorrect = 793,
        ywToolCutomer = 797,
        ywToolShowSign = 5756,
        ywToolOptions = 522,
        ywTableDrawTable = 2059,
        ywTableInsert = 30444,
        ywTableInsertTable = 3392,
        ywTableColumnsInsertLeft = 3685,
        ywTableColumnsInsertRight = 3687,
        ywTableRowsInsertAbove = 3681,
        ywTableRowsInsertBelow = 3683,
        ywTableInsertCell = 295,
        ywTableDelete = 30445,
        ywTableDeleteTable = 3679,
        ywTableDeleteColumn = 294,
        ywTableDeleteRow = 293,
        ywTableDeleteCell = 292,
        ywTableSelect = 30446,
        ywTableSelectTable = 803,
        ywTableSelectColumn = 802,
        ywTableSelectRow = 801,
        ywTableSelectCell = 3678,
        ywTableMergeCell = 798,
        ywTableSplitCell = 800,
        ywTableSplitTable = 808,
        ywTableAutoFormatStyle = 6400,
        ywTableAutoFit = 30460,
        ywTableAutoFitContents = 3907,
        ywTableAutoFitWindow = 3908,
        ywTableAutoFitFixedColumnWidth = 3909,
        ywTableRowsDistribute = 2068,
        ywTableColumnsDistribute = 2067,
        ywTableRepeatHeaderRows = 805,
        ywTableInsertMultidiagonalCell = 3961,
        ywTableTranslate = 30447,
        ywTableConvertTextToTable = 806,
        ywTableConvertTableToText = 929,
        ywTableSortDialogClassic = 928,
        ywTableFormula = 807,
        ywTableHideVoid = 2816,
        ywTableProperties = 3704,
        ywWindowNew = 303,
        ywWindowArrangeAll = 298,
        ywWindowSideBySide = 7698,
        ywWindowSplit = 302,
        ywStandardDocumentNew = 2520,
        ywStandardOpen = 23,
        ywStandardSave = 3,
        ywStandardPermission = 9004,
        ywStandardSendMailRecipient = 3738,
        ywStandardPrint = 2521,
        ywStandardPrintPreview = 109,
        ywStandardChineseTranslation = 4026,
        ywStandardSpellingAndGrammar = 2566,
        ywStandardResearch = 7343,
        ywStandardCut = 21,
        ywStandardCopy = 19,
        ywStandardPaste = 22,
        ywStandardFormatPainter = 108,
        ywStandardUndo = 128,
        ywStandardRedo = 129,
        ywStandardInsertInkComment = 9404,
        ywStandardInsertHyperlink = 1576,
        ywStandardTableAndBorderToolbar = 916,
        ywStandardInsertTable = 333,
        ywStandardInsertWorksheet = 142,
        ywStandardColumns = 9,
        ywStandardTextDirection = 2872,
        ywStandardDrawing = 204,
        ywStandardDocumentMap = 1714,
        ywStandardEditingMarks = 3900,
        ywStandardZoom = 1733,
        ywStandardHelp = 984,
        ywStandardReading = 7226,
        ywStandardPrintDialog = 4,
        ywStandardClose = 106,
        ywStandardEnvelopesAndLabels = 794,
        ywStandardFindDialog = 141,
        ywFormattingStylesPane = 5757,
        ywFormattingStyleGallery = 1732,
        ywFormattingFont = 1728,
        ywFormattingFontSize = 1731,
        ywFormattingBold = 113,
        ywFormattingItalic = 114,
        ywFormattingUnderline = 3962,
        ywFormattingCharacterBorder = 3517,
        ywFormattingCharacterShading = 3518,
        ywFormattingCharacterScaling = 386,
        ywFormattingAlignJustify = 123,
        ywFormattingAlignCenter = 122,
        ywFormattingAlignRight = 121,
        ywFormattingParagraphDistributed = 2792,
        ywFormattingLineSpace = 5734,
        ywFormattingTextDirectionLeftToRight = 1846,
        ywFormattingTextDirectionRightToLeft = 1847,
        ywFormattingNumbering = 11,
        ywFormattingBullets = 12,
        ywFormattingIndentDecreaseWord = 3473,
        ywFormattingIndentIncreaseWord = 3472,
        ywFormattingHighLight = 340,
        ywFormattingFontColor = 401,
        ywFormattingPhoneticGuide = 3511,
        ywFormattingCharactersEnclose = 3969,
        ywFormattingFontSizeDecreaseWord = 63,
        ywFormattingFontSizeIncreaseWord = 62,
        ywFormattingSuperscript = 57,
        ywFormattingSubscript = 58,
        ywFormattingLanguage = 3745,
        ywTrack = 2041

    }
    export const enum YwMenuID {
        ywFile = 30002,
        ywEdit = 30003,
        ywView = 30004,
        ywInsert = 30005,
        ywFormat = 30006,
        ywTool = 30007,
        ywTable = 30008,
        ywWindow = 30009,
        ywHelp = 30010,
        ywFileNew = 18,
        ywFileOpen = 23,
        ywFileClose = 106,
        ywFileSave = 3,
        ywFileSaveAs = 748,
        ywFileSaveAsHtml = 3823,
        ywFileSaveWorkspace = 846,
        ywFileFind = 5905,
        ywFilePermissionView = 7994,
        ywFileVersionHistory = 2522,
        ywFileWebPagePreview = 3655,
        ywFilePageSetupDialog = 247,
        ywFilePreviewPrint = 109,
        ywFilePrint = 4,
        ywFileProperties = 750,
        ywFileExit = 752,
        ywEditUndo = 128,
        ywEditRedo = 335,
        ywEditCut = 21,
        ywEditCopy = 19,
        ywEditPaste = 22,
        ywEditPasteSpecialDialog = 755,
        ywEditPasteAsHyperlink = 2787,
        ywEditClearMenuWord = 30021,
        ywEditClearFormat = 872,
        ywEditClearContent = 7024,
        ywEditSelectAll = 756,
        ywEditFind = 141,
        ywEditReplace = 313,
        ywEditGoto = 757,
        ywEditLinksToFiles = 759,
        ywViewNormalView = 224,
        ywViewWebLayoutView = 3621,
        ywViewPrintLayoutView = 287,
        ywViewFullScreenReadingView = 9377,
        ywViewMasterDocumentView = 4387,
        ywViewTaskPane = 5746,
        ywViewToolBars = 30045,
        ywViewRuler = 7391,
        ywViewShowParasTag = 3252,
        ywViewGridlines = 300,
        ywViewDocumentMap = 1714,
        ywViewHeaderAndFooter = 762,
        ywViewFootnotesEndnotes = 764,
        ywViewSign = 6547,
        ywViewHtmlSource = 3902,
        ywViewFullScreen = 178,
        ywViewZoom = 925,
        ywFormatFont = 253,
        ywFormatParagraph = 779,
        ywFormatBulletsAndNumbers = 783,
        ywFormatBordersShading = 781,
        ywFormatColumns = 9,
        ywFormatTabs = 780,
        ywFormatDropCap = 289,
        ywFormatTextDirection = 782,
        ywFormatChangeCase = 309,
        ywFormatFitText = 3510,
        ywFormatChineseLayout = 30463,
        ywFormatAsianLayoutPhoneticGuide = 3511,
        ywFormatAsianLayoutCharactersEnclose = 3969,
        ywFormatAsianLayoutHorizontalInVertical = 3963,
        ywFormatAsianLayoutCombineCharacters = 3512,
        ywFormatAsianLayoutTwoLinesInOne = 3964,
        ywFormatBackground = 30403,
        ywFormatTheme = 3623,
        ywFormatFrameset = 30452,
        ywFormatAutoFormat = 144,
        ywFormatStylesPane = 5757,
        ywFormatRevealFormatting = 6094,
        ywFormatObject = 2327,
        ywToolSpellingAndGrammar = 2566,
        ywToolResearchPane = 7343,
        ywToolLanguage = 30182,
        ywToolSetLanguage = 790,
        ywToolChineseTranslation = 3958,
        ywToolTranslation = 6111,
        ywToolThesaurus = 9056,
        ywToolHyphenation = 2800,
        ywToolWordCount = 792,
        ywToolAutoSummarize = 1709,
        ywToolPhonetic = 5764,
        ywToolShareWorkspace = 7710,
        ywToolTrackChanges = 2041,
        ywToolCompareAndCombine = 2338,
        ywToolProtect = 7116,
        ywToolWorkOnline = 30468,
        ywToolMeetingStart = 3727,
        ywToolMeetingArrange = 4179,
        ywToolWebDiscuss = 4177,
        ywToolLetterAndMail = 31171,
        ywToolMailMerge = 6070,
        ywToolShowMailMergeToolbar = 5911,
        ywToolEnvelopesAndLabels = 794,
        ywToolChineseEnvelopeWizard = 3972,
        ywToolEnglishEnvelopeWizard = 796,
        ywToolMacro = 30017,
        ywToolMacroDialog = 186,
        ywToolRecordMacro = 2780,
        ywToolMacroSecurity = 3627,
        ywToolVisualBasic = 1695,
        ywToolMSScriptEditor = 3631,
        ywToolComAddIns = 3754,
        ywToolDocumentTemplate = 751,
        ywToolAutoCorrect = 793,
        ywToolCutomer = 797,
        ywToolShowSign = 5756,
        ywToolOptions = 522,
        ywTableDrawTable = 2059,
        ywTableInsert = 30444,
        ywTableInsertTable = 3392,
        ywTableColumnsInsertLeft = 3685,
        ywTableColumnsInsertRight = 3687,
        ywTableRowsInsertAbove = 3681,
        ywTableRowsInsertBelow = 3683,
        ywTableInsertCell = 295,
        ywTableDelete = 30445,
        ywTableDeleteTable = 3679,
        ywTableDeleteColumn = 294,
        ywTableDeleteRow = 293,
        ywTableDeleteCell = 292,
        ywTableSelect = 30446,
        ywTableSelectTable = 803,
        ywTableSelectColumn = 802,
        ywTableSelectRow = 801,
        ywTableSelectCell = 3678,
        ywTableMergeCell = 798,
        ywTableSplitCell = 800,
        ywTableSplitTable = 808,
        ywTableAutoFormatStyle = 6400,
        ywTableAutoFit = 30460,
        ywTableAutoFitContents = 3907,
        ywTableAutoFitWindow = 3908,
        ywTableAutoFitFixedColumnWidth = 3909,
        ywTableRowsDistribute = 2068,
        ywTableColumnsDistribute = 2067,
        ywTableRepeatHeaderRows = 805,
        ywTableInsertMultidiagonalCell = 3961,
        ywTableTranslate = 30447,
        ywTableConvertTextToTable = 806,
        ywTableConvertTableToText = 929,
        ywTableSortDialogClassic = 928,
        ywTableFormula = 807,
        ywTableHideVoid = 2816,
        ywTableProperties = 3704,
        ywWindowNew = 303,
        ywWindowArrangeAll = 298,
        ywWindowSideBySide = 7698,
        ywWindowSplit = 302,
        ywStandardDocumentNew = 2520,
        ywStandardOpen = 23,
        ywStandardSave = 3,
        ywStandardPermission = 9004,
        ywStandardSendMailRecipient = 3738,
        ywStandardPrint = 2521,
        ywStandardPrintPreview = 109,
        ywStandardChineseTranslation = 4026,
        ywStandardSpellingAndGrammar = 2566,
        ywStandardResearch = 7343,
        ywStandardCut = 21,
        ywStandardCopy = 19,
        ywStandardPaste = 22,
        ywStandardFormatPainter = 108,
        ywStandardUndo = 128,
        ywStandardRedo = 129,
        ywStandardInsertInkComment = 9404,
        ywStandardInsertHyperlink = 1576,
        ywStandardTableAndBorderToolbar = 916,
        ywStandardInsertTable = 333,
        ywStandardInsertWorksheet = 142,
        ywStandardColumns = 9,
        ywStandardTextDirection = 2872,
        ywStandardDrawing = 204,
        ywStandardDocumentMap = 1714,
        ywStandardEditingMarks = 3900,
        ywStandardZoom = 1733,
        ywStandardHelp = 984,
        ywStandardReading = 7226,
        ywStandardPrintDialog = 4,
        ywStandardClose = 106,
        ywStandardEnvelopesAndLabels = 794,
        ywStandardFindDialog = 141,
        ywFormattingStylesPane = 5757,
        ywFormattingStyleGallery = 1732,
        ywFormattingFont = 1728,
        ywFormattingFontSize = 1731,
        ywFormattingBold = 113,
        ywFormattingItalic = 114,
        ywFormattingUnderline = 3962,
        ywFormattingCharacterBorder = 3517,
        ywFormattingCharacterShading = 3518,
        ywFormattingCharacterScaling = 386,
        ywFormattingAlignJustify = 123,
        ywFormattingAlignCenter = 122,
        ywFormattingAlignRight = 121,
        ywFormattingParagraphDistributed = 2792,
        ywFormattingLineSpace = 5734,
        ywFormattingTextDirectionLeftToRight = 1846,
        ywFormattingTextDirectionRightToLeft = 1847,
        ywFormattingNumbering = 11,
        ywFormattingBullets = 12,
        ywFormattingIndentDecreaseWord = 3473,
        ywFormattingIndentIncreaseWord = 3472,
        ywFormattingHighLight = 340,
        ywFormattingFontColor = 401,
        ywFormattingPhoneticGuide = 3511,
        ywFormattingCharactersEnclose = 3969,
        ywFormattingFontSizeDecreaseWord = 63,
        ywFormattingFontSizeIncreaseWord = 62,
        ywFormattingSuperscript = 57,
        ywFormattingSubscript = 58,
        ywFormattingLanguage = 3745,
        ywTrack = 2041

    }
    export const enum YwMergeTarget {
        ywMergeTargetCurrent = 1,
        ywMergeTargetNew = 2,
        ywMergeTargetSelected = 0

    }
    export const enum YwMergeTarget {
        ywMergeTargetCurrent = 1,
        ywMergeTargetNew = 2,
        ywMergeTargetSelected = 0

    }
    export const enum YwMonthNames {
        ywMonthNamesArabic = 0,
        ywMonthNamesEnglish = 1,
        ywMonthNamesFrench = 2

    }
    export const enum YwMonthNames {
        ywMonthNamesArabic = 0,
        ywMonthNamesEnglish = 1,
        ywMonthNamesFrench = 2

    }
    export const enum YwMovementType {
        ywMove = 0,
        ywExtend = 1

    }
    export const enum YwMovementType {
        ywMove = 0,
        ywExtend = 1

    }
    export const enum YwMultipleWordConversionsMode {
        ywHangulToHanja = 0,
        ywHanjaToHangul = 1

    }
    export const enum YwMultipleWordConversionsMode {
        ywHangulToHanja = 0,
        ywHanjaToHangul = 1

    }
    export const enum YwNewDocumentType {
        ywNewBlankDocument = 0,
        ywNewEmailMessage = 2,
        ywNewFrameset = 3,
        ywNewWebPage = 1,
        ywNewXMLDocument = 4

    }
    export const enum YwNewDocumentType {
        ywNewBlankDocument = 0,
        ywNewEmailMessage = 2,
        ywNewFrameset = 3,
        ywNewWebPage = 1,
        ywNewXMLDocument = 4

    }
    export const enum YwNumberingRule {
        ywRestartContinuous = 0,
        ywRestartSection = 1,
        ywRestartPage = 2

    }
    export const enum YwNumberingRule {
        ywRestartContinuous = 0,
        ywRestartSection = 1,
        ywRestartPage = 2

    }
    export const enum YwNumberStyle {
        ywNumberStyleArabic = 0,
        ywNumberStyleUppercaseRoman = 1,
        ywNumberStyleLowercaseRoman = 2,
        ywNumberStyleUppercaseLetter = 3,
        ywNumberStyleLowercaseLetter = 4,
        ywNumberStyleOrdinal = 5,
        ywNumberStyleCardinalText = 6,
        ywNumberStyleOrdinalText = 7,
        ywNoteNumberStyleSymbol = 9,
        ywNumberStyleKanji = 10,
        ywNumberStyleKanjiDigit = 11,
        ywNumberStyleAiueoHalfWidth = 12,
        ywNumberStyleIrohaHalfWidth = 13,
        ywNumberStyleArabicFullWidth = 14,
        ywNumberStyleKanjiTraditional = 16,
        ywNumberStyleKanjiTraditional2 = 17,
        ywNumberStyleNumberInCircle = 18,
        ywNumberStyleAiueo = 20,
        ywNumberStyleIroha = 21,
        ywNumberStyleArabicLZ = 22,
        ywNumberStyleBullet = 23,
        ywNumberStyleGanada = 24,
        ywNumberStyleChosung = 25,
        ywNumberStyleGBNum1 = 26,
        ywNumberStyleGBNum2 = 27,
        ywNumberStyleGBNum3 = 28,
        ywNumberStyleGBNum4 = 29,
        ywNumberStyleZodiac1 = 30,
        ywNumberStyleZodiac2 = 31,
        ywNumberStyleZodiac3 = 32,
        ywNumberStyleTradChinNum1 = 33,
        ywNumberStyleTradChinNum2 = 34,
        ywNumberStyleTradChinNum3 = 35,
        ywNumberStyleTradChinNum4 = 36,
        ywNumberStyleSimpChinNum1 = 37,
        ywNumberStyleSimpChinNum2 = 38,
        ywNumberStyleSimpChinNum3 = 39,
        ywNumberStyleSimpChinNum4 = 40,
        ywNumberStyleHanjaRead = 41,
        ywNumberStyleHanjaReadDigit = 42,
        ywNumberStyleHangul = 43,
        ywNumberStyleHanja = 44,
        ywNumberStyleHebrew1 = 45,
        ywNumberStyleArabic1 = 46,
        ywNumberStyleHebrew2 = 47,
        ywNumberStyleArabic2 = 48,
        ywNumberStyleHindiLetter1 = 49,
        ywNumberStyleHindiLetter2 = 50,
        ywNumberStyleHindiArabic = 51,
        ywNumberStyleHindiCardinalText = 52,
        ywNumberStyleThaiLetter = 53,
        ywNumberStyleThaiArabic = 54,
        ywNumberStyleThaiCardinalText = 55,
        ywNumberStyleVietCardinalText = 56,
        ywNumberStyleNumberInDash = 57,
        ywNumberStyleLowercaseRussian = 58,
        ywNumberStyleUppercaseRussian = 59,
        ywNumberStylePictureBullet = 249,
        ywNumberStyleLegal = 253,
        ywNumberStyleLegalLZ = 254,
        ywNumberStyleNone = 255

    }
    export const enum YwNumberStyle {
        ywNumberStyleArabic = 0,
        ywNumberStyleUppercaseRoman = 1,
        ywNumberStyleLowercaseRoman = 2,
        ywNumberStyleUppercaseLetter = 3,
        ywNumberStyleLowercaseLetter = 4,
        ywNumberStyleOrdinal = 5,
        ywNumberStyleCardinalText = 6,
        ywNumberStyleOrdinalText = 7,
        ywNoteNumberStyleSymbol = 9,
        ywNumberStyleKanji = 10,
        ywNumberStyleKanjiDigit = 11,
        ywNumberStyleAiueoHalfWidth = 12,
        ywNumberStyleIrohaHalfWidth = 13,
        ywNumberStyleArabicFullWidth = 14,
        ywNumberStyleKanjiTraditional = 16,
        ywNumberStyleKanjiTraditional2 = 17,
        ywNumberStyleNumberInCircle = 18,
        ywNumberStyleAiueo = 20,
        ywNumberStyleIroha = 21,
        ywNumberStyleArabicLZ = 22,
        ywNumberStyleBullet = 23,
        ywNumberStyleGanada = 24,
        ywNumberStyleChosung = 25,
        ywNumberStyleGBNum1 = 26,
        ywNumberStyleGBNum2 = 27,
        ywNumberStyleGBNum3 = 28,
        ywNumberStyleGBNum4 = 29,
        ywNumberStyleZodiac1 = 30,
        ywNumberStyleZodiac2 = 31,
        ywNumberStyleZodiac3 = 32,
        ywNumberStyleTradChinNum1 = 33,
        ywNumberStyleTradChinNum2 = 34,
        ywNumberStyleTradChinNum3 = 35,
        ywNumberStyleTradChinNum4 = 36,
        ywNumberStyleSimpChinNum1 = 37,
        ywNumberStyleSimpChinNum2 = 38,
        ywNumberStyleSimpChinNum3 = 39,
        ywNumberStyleSimpChinNum4 = 40,
        ywNumberStyleHanjaRead = 41,
        ywNumberStyleHanjaReadDigit = 42,
        ywNumberStyleHangul = 43,
        ywNumberStyleHanja = 44,
        ywNumberStyleHebrew1 = 45,
        ywNumberStyleArabic1 = 46,
        ywNumberStyleHebrew2 = 47,
        ywNumberStyleArabic2 = 48,
        ywNumberStyleHindiLetter1 = 49,
        ywNumberStyleHindiLetter2 = 50,
        ywNumberStyleHindiArabic = 51,
        ywNumberStyleHindiCardinalText = 52,
        ywNumberStyleThaiLetter = 53,
        ywNumberStyleThaiArabic = 54,
        ywNumberStyleThaiCardinalText = 55,
        ywNumberStyleVietCardinalText = 56,
        ywNumberStyleNumberInDash = 57,
        ywNumberStyleLowercaseRussian = 58,
        ywNumberStyleUppercaseRussian = 59,
        ywNumberStylePictureBullet = 249,
        ywNumberStyleLegal = 253,
        ywNumberStyleLegalLZ = 254,
        ywNumberStyleNone = 255

    }
    export const enum YwNumberType {
        ywNumberParagraph = 1,
        ywNumberListNum = 2,
        ywNumberAllNumbers = 3

    }
    export const enum YwNumberType {
        ywNumberParagraph = 1,
        ywNumberListNum = 2,
        ywNumberAllNumbers = 3

    }
    export const enum YwOLEVerb {
        ywOLEVerbDiscardUndoState = -6,
        ywOLEVerbInPlaceActivate = -5,
        ywOLEVerbUIActivate = -4,
        ywOLEVerbHide = -3,
        ywOLEVerbOpen = -2,
        ywOLEVerbShow = -1,
        ywOLEVerbPrimary = 0

    }
    export const enum YwOLEVerb {
        ywOLEVerbDiscardUndoState = -6,
        ywOLEVerbInPlaceActivate = -5,
        ywOLEVerbUIActivate = -4,
        ywOLEVerbHide = -3,
        ywOLEVerbOpen = -2,
        ywOLEVerbShow = -1,
        ywOLEVerbPrimary = 0

    }
    export const enum YwOpenFormat {
        ywOpenFormatAllWord = 6,
        ywOpenFormatAuto = 0,
        ywOpenFormatDocument = 1,
        ywOpenFormatEncodedText = 5,
        ywOpenFormatRTF = 3,
        ywOpenFormatTemplate = 2,
        ywOpenFormatText = 4,
        ywOpenFormatUnicodeText = 5,
        ywOpenFormatWebPages = 7,
        ywOpenFormatXML = 8

    }
    export const enum YwOpenFormat {
        ywOpenFormatAllWord = 6,
        ywOpenFormatAuto = 0,
        ywOpenFormatDocument = 1,
        ywOpenFormatEncodedText = 5,
        ywOpenFormatRTF = 3,
        ywOpenFormatTemplate = 2,
        ywOpenFormatText = 4,
        ywOpenFormatUnicodeText = 5,
        ywOpenFormatWebPages = 7,
        ywOpenFormatXML = 8

    }
    export const enum YwOrientation {
        ywOrientLandscape = 1,
        ywOrientPortrait = 0

    }
    export const enum YwOrientation {
        ywOrientLandscape = 1,
        ywOrientPortrait = 0

    }
    export const enum YwOriginalFormat {
        ywWordDocument = 0,
        ywOriginalDocumentFormat = 1,
        ywPromptUser = 2

    }
    export const enum YwOriginalFormat {
        ywWordDocument = 0,
        ywOriginalDocumentFormat = 1,
        ywPromptUser = 2

    }
    export const enum YwoTriState {
        ywoTrue = -1,
        ywoFalse = 0

    }
    export const enum YwoTriState {
        ywoTrue = -1,
        ywoFalse = 0

    }
    export const enum YwOutlineLevel {
        wdOutlineLevel1 = 1,
        wdOutlineLevel2 = 2,
        wdOutlineLevel3 = 3,
        wdOutlineLevel4 = 4,
        wdOutlineLevel5 = 5,
        wdOutlineLevel6 = 6,
        wdOutlineLevel7 = 7,
        wdOutlineLevel8 = 8,
        wdOutlineLevel9 = 9,
        wdOutlineLevelBodyText = 10

    }
    export const enum YwOutlineLevel {
        wdOutlineLevel1 = 1,
        wdOutlineLevel2 = 2,
        wdOutlineLevel3 = 3,
        wdOutlineLevel4 = 4,
        wdOutlineLevel5 = 5,
        wdOutlineLevel6 = 6,
        wdOutlineLevel7 = 7,
        wdOutlineLevel8 = 8,
        wdOutlineLevel9 = 9,
        wdOutlineLevelBodyText = 10

    }
    export const enum YwPageBorderArt {
        ywArtApples = 1,
        ywArtArchedScallops = 97,
        ywArtBabyPacifier = 70,
        ywArtBabyRattle = 71,
        ywArtBalloons3Colors = 11,
        ywArtBalloonsHotAir = 12,
        ywArtBasicBlackDashes = 155,
        ywArtBasicBlackDots = 156,
        ywArtBasicBlackSquares = 154,
        ywArtBasicThinLines = 151,
        ywArtBasicWhiteDashes = 152,
        ywArtBasicWhiteDots = 147,
        ywArtBasicWhiteSquares = 153,
        ywArtBasicWideInline = 150,
        ywArtBasicWideMidline = 148,
        ywArtBasicWideOutline = 149,
        ywArtBats = 37,
        ywArtBirds = 102,
        ywArtBirdsFlight = 35,
        ywArtCabins = 72,
        ywArtCakeSlice = 3,
        ywArtCandyCorn = 4,
        ywArtCelticKnotwork = 99,
        ywArtCertificateBanner = 158,
        ywArtChainLink = 128,
        ywArtChampagneBottle = 6,
        ywArtCheckedBarBlack = 145,
        ywArtCheckedBarColor = 61,
        ywArtCheckered = 144,
        ywArtChristmasTree = 8,
        ywArtCirclesLines = 91,
        ywArtCirclesRectangles = 140,
        ywArtClassicalWave = 56,
        ywArtClocks = 27,
        ywArtCompass = 54,
        ywArtConfetti = 31,
        ywArtConfettiGrays = 115,
        ywArtConfettiOutline = 116,
        ywArtConfettiStreamers = 14,
        ywArtConfettiWhite = 117,
        ywArtCornerTriangles = 141,
        ywArtCouponCutoutDashes = 163,
        ywArtCouponCutoutDots = 164,
        ywArtCrazyMaze = 100,
        ywArtCreaturesButterfly = 32,
        ywArtCreaturesFish = 34,
        ywArtCreaturesInsects = 142,
        ywArtCreaturesLadyBug = 33,
        ywArtCrossStitch = 138,
        ywArtCup = 67,
        ywArtDecoArch = 89,
        ywArtDecoArchColor = 50,
        ywArtDecoBlocks = 90,
        ywArtDiamondsGray = 88,
        ywArtDoubleD = 55,
        ywArtDoubleDiamonds = 127,
        ywArtEarth1 = 22,
        ywArtEarth2 = 21,
        ywArtEclipsingSquares1 = 101,
        ywArtEclipsingSquares2 = 86,
        ywArtEggsBlack = 66,
        ywArtFans = 51,
        ywArtFilm = 52,
        ywArtFirecrackers = 28,
        ywArtFlowersBlockPrint = 49,
        ywArtFlowersDaisies = 48,
        ywArtFlowersModern1 = 45,
        ywArtFlowersModern2 = 44,
        ywArtFlowersPansy = 43,
        ywArtFlowersRedRose = 39,
        ywArtFlowersRoses = 38,
        ywArtFlowersTeacup = 103,
        ywArtFlowersTiny = 42,
        ywArtGems = 139,
        ywArtGingerbreadMan = 69,
        ywArtGradient = 122,
        ywArtHandmade1 = 159,
        ywArtHandmade2 = 160,
        ywArtHeartBalloon = 16,
        ywArtHeartGray = 68,
        ywArtHearts = 15,
        ywArtHeebieJeebies = 120,
        ywArtHolly = 41,
        ywArtHouseFunky = 73,
        ywArtHypnotic = 87,
        ywArtIceCreamCones = 5,
        ywArtLightBulb = 121,
        ywArtLightning1 = 53,
        ywArtLightning2 = 119,
        ywArtMapleLeaf = 81,
        ywArtMapleMuffins = 2,
        ywArtMapPins = 30,
        ywArtMarquee = 146,
        ywArtMarqueeToothed = 131,
        ywArtMoons = 125,
        ywArtMosaic = 118,
        ywArtMusicNotes = 79,
        ywArtNorthwest = 104,
        ywArtOvals = 126,
        ywArtPackages = 26,
        ywArtPalmsBlack = 80,
        ywArtPalmsColor = 10,
        ywArtPaperClips = 82,
        ywArtPapyrus = 92,
        ywArtPartyFavor = 13,
        ywArtPartyGlass = 7,
        ywArtPencils = 25,
        ywArtPeople = 84,
        ywArtPeopleHats = 23,
        ywArtPeopleWaving = 85,
        ywArtPoinsettias = 40,
        ywArtPostageStamp = 135,
        ywArtPumpkin1 = 65,
        ywArtPushPinNote1 = 63,
        ywArtPushPinNote2 = 64,
        ywArtPyramids = 113,
        ywArtPyramidsAbove = 114,
        ywArtQuadrants = 60,
        ywArtRings = 29,
        ywArtSafari = 98,
        ywArtSawtooth = 133,
        ywArtSawtoothGray = 134,
        ywArtScaredCat = 36,
        ywArtSeattle = 78,
        ywArtShadowedSquares = 57,
        ywArtSharksTeeth = 132,
        ywArtShorebirdTracks = 83,
        ywArtSkyrocket = 77,
        ywArtSnowflakeFancy = 76,
        ywArtSnowflakes = 75,
        ywArtSombrero = 24,
        ywArtSouthwest = 105,
        ywArtStars = 19,
        ywArtStars3D = 17,
        ywArtStarsBlack = 74,
        ywArtStarsShadowed = 18,
        ywArtStarsTop = 157,
        ywArtSun = 20,
        ywArtSwirligig = 62,
        ywArtTornPaper = 161,
        ywArtTornPaperBlack = 162,
        ywArtTrees = 9,
        ywArtTriangleParty = 123,
        ywArtTriangles = 129,
        ywArtTribal1 = 130,
        ywArtTribal2 = 109,
        ywArtTribal3 = 108,
        ywArtTribal4 = 107,
        ywArtTribal5 = 110,
        ywArtTribal6 = 106,
        ywArtTwistedLines1 = 58,
        ywArtTwistedLines2 = 124,
        ywArtVine = 47,
        ywArtWaveline = 59,
        ywArtWeavingAngles = 96,
        ywArtWeavingBraid = 94,
        ywArtWeavingRibbon = 95,
        ywArtWeavingStrips = 136,
        ywArtWhiteFlowers = 46,
        ywArtWoodwork = 93,
        ywArtXIllusions = 111,
        ywArtZanyTriangles = 112,
        ywArtZigZag = 137,
        ywArtZigZagStitch = 143

    }
    export const enum YwPageBorderArt {
        ywArtApples = 1,
        ywArtArchedScallops = 97,
        ywArtBabyPacifier = 70,
        ywArtBabyRattle = 71,
        ywArtBalloons3Colors = 11,
        ywArtBalloonsHotAir = 12,
        ywArtBasicBlackDashes = 155,
        ywArtBasicBlackDots = 156,
        ywArtBasicBlackSquares = 154,
        ywArtBasicThinLines = 151,
        ywArtBasicWhiteDashes = 152,
        ywArtBasicWhiteDots = 147,
        ywArtBasicWhiteSquares = 153,
        ywArtBasicWideInline = 150,
        ywArtBasicWideMidline = 148,
        ywArtBasicWideOutline = 149,
        ywArtBats = 37,
        ywArtBirds = 102,
        ywArtBirdsFlight = 35,
        ywArtCabins = 72,
        ywArtCakeSlice = 3,
        ywArtCandyCorn = 4,
        ywArtCelticKnotwork = 99,
        ywArtCertificateBanner = 158,
        ywArtChainLink = 128,
        ywArtChampagneBottle = 6,
        ywArtCheckedBarBlack = 145,
        ywArtCheckedBarColor = 61,
        ywArtCheckered = 144,
        ywArtChristmasTree = 8,
        ywArtCirclesLines = 91,
        ywArtCirclesRectangles = 140,
        ywArtClassicalWave = 56,
        ywArtClocks = 27,
        ywArtCompass = 54,
        ywArtConfetti = 31,
        ywArtConfettiGrays = 115,
        ywArtConfettiOutline = 116,
        ywArtConfettiStreamers = 14,
        ywArtConfettiWhite = 117,
        ywArtCornerTriangles = 141,
        ywArtCouponCutoutDashes = 163,
        ywArtCouponCutoutDots = 164,
        ywArtCrazyMaze = 100,
        ywArtCreaturesButterfly = 32,
        ywArtCreaturesFish = 34,
        ywArtCreaturesInsects = 142,
        ywArtCreaturesLadyBug = 33,
        ywArtCrossStitch = 138,
        ywArtCup = 67,
        ywArtDecoArch = 89,
        ywArtDecoArchColor = 50,
        ywArtDecoBlocks = 90,
        ywArtDiamondsGray = 88,
        ywArtDoubleD = 55,
        ywArtDoubleDiamonds = 127,
        ywArtEarth1 = 22,
        ywArtEarth2 = 21,
        ywArtEclipsingSquares1 = 101,
        ywArtEclipsingSquares2 = 86,
        ywArtEggsBlack = 66,
        ywArtFans = 51,
        ywArtFilm = 52,
        ywArtFirecrackers = 28,
        ywArtFlowersBlockPrint = 49,
        ywArtFlowersDaisies = 48,
        ywArtFlowersModern1 = 45,
        ywArtFlowersModern2 = 44,
        ywArtFlowersPansy = 43,
        ywArtFlowersRedRose = 39,
        ywArtFlowersRoses = 38,
        ywArtFlowersTeacup = 103,
        ywArtFlowersTiny = 42,
        ywArtGems = 139,
        ywArtGingerbreadMan = 69,
        ywArtGradient = 122,
        ywArtHandmade1 = 159,
        ywArtHandmade2 = 160,
        ywArtHeartBalloon = 16,
        ywArtHeartGray = 68,
        ywArtHearts = 15,
        ywArtHeebieJeebies = 120,
        ywArtHolly = 41,
        ywArtHouseFunky = 73,
        ywArtHypnotic = 87,
        ywArtIceCreamCones = 5,
        ywArtLightBulb = 121,
        ywArtLightning1 = 53,
        ywArtLightning2 = 119,
        ywArtMapleLeaf = 81,
        ywArtMapleMuffins = 2,
        ywArtMapPins = 30,
        ywArtMarquee = 146,
        ywArtMarqueeToothed = 131,
        ywArtMoons = 125,
        ywArtMosaic = 118,
        ywArtMusicNotes = 79,
        ywArtNorthwest = 104,
        ywArtOvals = 126,
        ywArtPackages = 26,
        ywArtPalmsBlack = 80,
        ywArtPalmsColor = 10,
        ywArtPaperClips = 82,
        ywArtPapyrus = 92,
        ywArtPartyFavor = 13,
        ywArtPartyGlass = 7,
        ywArtPencils = 25,
        ywArtPeople = 84,
        ywArtPeopleHats = 23,
        ywArtPeopleWaving = 85,
        ywArtPoinsettias = 40,
        ywArtPostageStamp = 135,
        ywArtPumpkin1 = 65,
        ywArtPushPinNote1 = 63,
        ywArtPushPinNote2 = 64,
        ywArtPyramids = 113,
        ywArtPyramidsAbove = 114,
        ywArtQuadrants = 60,
        ywArtRings = 29,
        ywArtSafari = 98,
        ywArtSawtooth = 133,
        ywArtSawtoothGray = 134,
        ywArtScaredCat = 36,
        ywArtSeattle = 78,
        ywArtShadowedSquares = 57,
        ywArtSharksTeeth = 132,
        ywArtShorebirdTracks = 83,
        ywArtSkyrocket = 77,
        ywArtSnowflakeFancy = 76,
        ywArtSnowflakes = 75,
        ywArtSombrero = 24,
        ywArtSouthwest = 105,
        ywArtStars = 19,
        ywArtStars3D = 17,
        ywArtStarsBlack = 74,
        ywArtStarsShadowed = 18,
        ywArtStarsTop = 157,
        ywArtSun = 20,
        ywArtSwirligig = 62,
        ywArtTornPaper = 161,
        ywArtTornPaperBlack = 162,
        ywArtTrees = 9,
        ywArtTriangleParty = 123,
        ywArtTriangles = 129,
        ywArtTribal1 = 130,
        ywArtTribal2 = 109,
        ywArtTribal3 = 108,
        ywArtTribal4 = 107,
        ywArtTribal5 = 110,
        ywArtTribal6 = 106,
        ywArtTwistedLines1 = 58,
        ywArtTwistedLines2 = 124,
        ywArtVine = 47,
        ywArtWaveline = 59,
        ywArtWeavingAngles = 96,
        ywArtWeavingBraid = 94,
        ywArtWeavingRibbon = 95,
        ywArtWeavingStrips = 136,
        ywArtWhiteFlowers = 46,
        ywArtWoodwork = 93,
        ywArtXIllusions = 111,
        ywArtZanyTriangles = 112,
        ywArtZigZag = 137,
        ywArtZigZagStitch = 143

    }
    export const enum YwPageFit {
        ywPageFitBestFit = 2,
        ywPageFitFullPage = 1,
        ywPageFitNone = 0,
        ywPageFitTextFit = 3

    }
    export const enum YwPageFit {
        ywPageFitBestFit = 2,
        ywPageFitFullPage = 1,
        ywPageFitNone = 0,
        ywPageFitTextFit = 3

    }
    export const enum YwPageNumberAlignment {
        ywAlignPageNumberLeft = 0,
        ywAlignPageNumberCenter = 1,
        ywAlignPageNumberRight = 2,
        ywAlignPageNumberInside = 3,
        ywAlignPageNumberOutside = 4

    }
    export const enum YwPageNumberAlignment {
        ywAlignPageNumberLeft = 0,
        ywAlignPageNumberCenter = 1,
        ywAlignPageNumberRight = 2,
        ywAlignPageNumberInside = 3,
        ywAlignPageNumberOutside = 4

    }
    export const enum YwPaperSize {
        ywPaper10x14 = 0,
        ywPaper11x17 = 1,
        ywPaperA3 = 6,
        ywPaperA4 = 7,
        ywPaperA4Small = 8,
        ywPaperA5 = 9,
        ywPaperB4 = 10,
        ywPaperB5 = 11,
        ywPaperCSheet = 12,
        ywPaperCustom = 41,
        ywPaperDSheet = 13,
        ywPaperEnvelope10 = 25,
        ywPaperEnvelope11 = 26,
        ywPaperEnvelope12 = 27,
        ywPaperEnvelope14 = 28,
        ywPaperEnvelope9 = 24,
        ywPaperEnvelopeB4 = 29,
        ywPaperEnvelopeB5 = 30,
        ywPaperEnvelopeB6 = 31,
        ywPaperEnvelopeC3 = 32,
        ywPaperEnvelopeC4 = 33,
        ywPaperEnvelopeC5 = 34,
        ywPaperEnvelopeC6 = 35,
        ywPaperEnvelopeC65 = 36,
        ywPaperEnvelopeDL = 37,
        ywPaperEnvelopeItaly = 38,
        ywPaperEnvelopeMonarch = 39,
        ywPaperEnvelopePersonal = 40,
        ywPaperESheet = 14,
        ywPaperExecutive = 5,
        ywPaperFanfoldLegalGerman = 15,
        ywPaperFanfoldStdGerman = 16,
        ywPaperFanfoldUS = 17,
        ywPaperFolio = 18,
        ywPaperLedger = 19,
        ywPaperLegal = 4,
        ywPaperLetter = 2,
        ywPaperLetterSmall = 3,
        ywPaperNote = 20,
        ywPaperQuarto = 21,
        ywPaperStatement = 22,
        ywPaperTabloid = 23

    }
    export const enum YwPaperSize {
        ywPaper10x14 = 0,
        ywPaper11x17 = 1,
        ywPaperA3 = 6,
        ywPaperA4 = 7,
        ywPaperA4Small = 8,
        ywPaperA5 = 9,
        ywPaperB4 = 10,
        ywPaperB5 = 11,
        ywPaperCSheet = 12,
        ywPaperCustom = 41,
        ywPaperDSheet = 13,
        ywPaperEnvelope10 = 25,
        ywPaperEnvelope11 = 26,
        ywPaperEnvelope12 = 27,
        ywPaperEnvelope14 = 28,
        ywPaperEnvelope9 = 24,
        ywPaperEnvelopeB4 = 29,
        ywPaperEnvelopeB5 = 30,
        ywPaperEnvelopeB6 = 31,
        ywPaperEnvelopeC3 = 32,
        ywPaperEnvelopeC4 = 33,
        ywPaperEnvelopeC5 = 34,
        ywPaperEnvelopeC6 = 35,
        ywPaperEnvelopeC65 = 36,
        ywPaperEnvelopeDL = 37,
        ywPaperEnvelopeItaly = 38,
        ywPaperEnvelopeMonarch = 39,
        ywPaperEnvelopePersonal = 40,
        ywPaperESheet = 14,
        ywPaperExecutive = 5,
        ywPaperFanfoldLegalGerman = 15,
        ywPaperFanfoldStdGerman = 16,
        ywPaperFanfoldUS = 17,
        ywPaperFolio = 18,
        ywPaperLedger = 19,
        ywPaperLegal = 4,
        ywPaperLetter = 2,
        ywPaperLetterSmall = 3,
        ywPaperNote = 20,
        ywPaperQuarto = 21,
        ywPaperStatement = 22,
        ywPaperTabloid = 23

    }
    export const enum YwPaperTray {
        ywPrinterAutomaticSheetFeed = 7,
        ywPrinterDefaultBin = 0,
        ywPrinterEnvelopeFeed = 5,
        ywPrinterFormSource = 15,
        ywPrinterLargeCapacityBin = 11,
        ywPrinterLargeFormatBin = 10,
        ywPrinterLowerBin = 2,
        ywPrinterManualEnvelopeFeed = 6,
        ywPrinterManualFeed = 4,
        ywPrinterMiddleBin = 3,
        ywPrinterOnlyBin = 1,
        ywPrinterPaperCassette = 14,
        ywPrinterSmallFormatBin = 9,
        ywPrinterTractorFeed = 8,
        ywPrinterUpperBin = 1

    }
    export const enum YwPaperTray {
        ywPrinterAutomaticSheetFeed = 7,
        ywPrinterDefaultBin = 0,
        ywPrinterEnvelopeFeed = 5,
        ywPrinterFormSource = 15,
        ywPrinterLargeCapacityBin = 11,
        ywPrinterLargeFormatBin = 10,
        ywPrinterLowerBin = 2,
        ywPrinterManualEnvelopeFeed = 6,
        ywPrinterManualFeed = 4,
        ywPrinterMiddleBin = 3,
        ywPrinterOnlyBin = 1,
        ywPrinterPaperCassette = 14,
        ywPrinterSmallFormatBin = 9,
        ywPrinterTractorFeed = 8,
        ywPrinterUpperBin = 1

    }
    export const enum YwParagraphAlignment {
        ywAlignParagraphCenter = 1,
        ywAlignParagraphDistribute = 4,
        ywAlignParagraphJustify = 3,
        ywAlignParagraphJustifyHi = 7,
        ywAlignParagraphJustifyLow = 8,
        ywAlignParagraphJustifyMed = 5,
        ywAlignParagraphLeft = 0,
        ywAlignParagraphRight = 2,
        ywAlignParagraphThaiJustify = 9

    }
    export const enum YwParagraphAlignment {
        ywAlignParagraphCenter = 1,
        ywAlignParagraphDistribute = 4,
        ywAlignParagraphJustify = 3,
        ywAlignParagraphJustifyHi = 7,
        ywAlignParagraphJustifyLow = 8,
        ywAlignParagraphJustifyMed = 5,
        ywAlignParagraphLeft = 0,
        ywAlignParagraphRight = 2,
        ywAlignParagraphThaiJustify = 9

    }
    export const enum YwPasteDataType {
        ywPasteBitmap = 4,
        ywPasteDeviceIndependentBitmap = 5,
        ywPasteEnhancedMetafile = 9,
        ywPasteHTML = 10,
        ywPasteHyperlink = 7,
        ywPasteMetafilePicture = 3,
        ywPasteOLEObject = 0,
        ywPasteRTF = 1,
        ywPasteShape = 8,
        ywPasteText = 2

    }
    export const enum YwPasteDataType {
        ywPasteBitmap = 4,
        ywPasteDeviceIndependentBitmap = 5,
        ywPasteEnhancedMetafile = 9,
        ywPasteHTML = 10,
        ywPasteHyperlink = 7,
        ywPasteMetafilePicture = 3,
        ywPasteOLEObject = 0,
        ywPasteRTF = 1,
        ywPasteShape = 8,
        ywPasteText = 2

    }
    export const enum YwPreferredWidthType {
        ywPreferredWidthAuto = 1,
        ywPreferredWidthPercent = 2,
        ywPreferredWidthPoints = 3

    }
    export const enum YwPreferredWidthType {
        ywPreferredWidthAuto = 1,
        ywPreferredWidthPercent = 2,
        ywPreferredWidthPoints = 3

    }
    export const enum YwPrintOutItem {
        ywPrintDocumentContent = 0,
        ywPrintProperties = 1,
        ywPrintDocumentWithMarkup = 2

    }
    export const enum YwPrintOutItem {
        ywPrintDocumentContent = 0,
        ywPrintProperties = 1,
        ywPrintDocumentWithMarkup = 2

    }
    export const enum YwPrintOutPages {
        ywPrintAllPages = 0,
        ywPrintOddPagesOnly = 1,
        ywPrintEvenPagesOnly = 2

    }
    export const enum YwPrintOutPages {
        ywPrintAllPages = 0,
        ywPrintOddPagesOnly = 1,
        ywPrintEvenPagesOnly = 2

    }
    export const enum YwPrintOutRange {
        ywPrintAllDocument = 0,
        ywPrintSelection = 1,
        ywPrintCurrentPage = 2,
        ywPrintFromTo = 3,
        ywPrintRangeOfPages = 4

    }
    export const enum YwPrintOutRange {
        ywPrintAllDocument = 0,
        ywPrintSelection = 1,
        ywPrintCurrentPage = 2,
        ywPrintFromTo = 3,
        ywPrintRangeOfPages = 4

    }
    export const enum YwProtectionType {
        ywNoProtection = -1,
        ywAllowOnlyRevisions = 0,
        ywAllowOnlyComments = 1,
        ywAllowOnlyFormFields = 2,
        ywAllowOnlyReading = 3

    }
    export const enum YwProtectionType {
        ywNoProtection = -1,
        ywAllowOnlyRevisions = 0,
        ywAllowOnlyComments = 1,
        ywAllowOnlyFormFields = 2,
        ywAllowOnlyReading = 3

    }
    export const enum YwRecoveryType {
        ywPasteDefault = 0,
        ywSingleCellText = 5,
        ywSingleCellTable = 6,
        ywListContinueNumbering = 7,
        ywListRestartNumbering = 8,
        ywTableAppendTable = 10,
        ywTableInsertAsRows = 11,
        ywTableOriginalFormatting = 12,
        ywChartPicture = 13,
        ywChart = 14,
        ywChartLinked = 15,
        ywFormatOriginalFormatting = 16,
        ywFormatSurroundingFormattingWithEmphasis = 20,
        ywFormatPlainText = 22,
        ywTableOverwriteCells = 23,
        ywListCombineWithExistingList = 24,
        ywListDontMerge = 25

    }
    export const enum YwRecoveryType {
        ywPasteDefault = 0,
        ywSingleCellText = 5,
        ywSingleCellTable = 6,
        ywListContinueNumbering = 7,
        ywListRestartNumbering = 8,
        ywTableAppendTable = 10,
        ywTableInsertAsRows = 11,
        ywTableOriginalFormatting = 12,
        ywChartPicture = 13,
        ywChart = 14,
        ywChartLinked = 15,
        ywFormatOriginalFormatting = 16,
        ywFormatSurroundingFormattingWithEmphasis = 20,
        ywFormatPlainText = 22,
        ywTableOverwriteCells = 23,
        ywListCombineWithExistingList = 24,
        ywListDontMerge = 25

    }
    export const enum YwRectangleType {
        ywLineBetweenColumnRectangle = 5,
        ywMarkupRectangle = 2,
        ywMarkupRectangleButton = 3,
        ywPageBorderRectangle = 4,
        ywSelection = 6,
        ywShapeRectangle = 1,
        ywSystem = 7,
        ywTextRectangle = 0

    }
    export const enum YwRectangleType {
        ywLineBetweenColumnRectangle = 5,
        ywMarkupRectangle = 2,
        ywMarkupRectangleButton = 3,
        ywPageBorderRectangle = 4,
        ywSelection = 6,
        ywShapeRectangle = 1,
        ywSystem = 7,
        ywTextRectangle = 0

    }
    export const enum YwReferenceKind {
        ywNumberFullContext = -4,
        ywNumberNoContext = -3,
        ywNumberRelativeContext = -2,
        ywContentText = -1,
        ywEntireCaption = 2,
        ywOnlyLabelAndNumber = 3,
        ywOnlyCaptionText = 4,
        ywFootnoteNumber = 5,
        ywEndnoteNumber = 6,
        ywPageNumber = 7,
        ywPosition = 15,
        ywFootnoteNumberFormatted = 16,
        ywEndnoteNumberFormatted = 17

    }
    export const enum YwReferenceKind {
        ywNumberFullContext = -4,
        ywNumberNoContext = -3,
        ywNumberRelativeContext = -2,
        ywContentText = -1,
        ywEntireCaption = 2,
        ywOnlyLabelAndNumber = 3,
        ywOnlyCaptionText = 4,
        ywFootnoteNumber = 5,
        ywEndnoteNumber = 6,
        ywPageNumber = 7,
        ywPosition = 15,
        ywFootnoteNumberFormatted = 16,
        ywEndnoteNumberFormatted = 17

    }
    export const enum YwReferenceType {
        ywRefTypeBookmark = 2,
        ywRefTypeEndnote = 4,
        ywRefTypeFootnote = 3,
        ywRefTypeHeading = 1,
        ywRefTypeNumberedItem = 0

    }
    export const enum YwReferenceType {
        ywRefTypeBookmark = 2,
        ywRefTypeEndnote = 4,
        ywRefTypeFootnote = 3,
        ywRefTypeHeading = 1,
        ywRefTypeNumberedItem = 0

    }
    export const enum YwRelativeHorizontalPosition {
        ywRelativeHorizontalPositionMargin = 0,
        ywRelativeHorizontalPositionPage = 1,
        ywRelativeHorizontalPositionColumn = 2,
        ywRelativeHorizontalPositionCharacter = 3,
        ywRelativeHorizontalPositionLeftMarginArea = 4,
        ywRelativeHorizontalPositionRightMarginArea = 5,
        ywRelativeHorizontalPositionInnerMarginArea = 6,
        ywRelativeHorizontalPositionOuterMarginArea = 7

    }
    export const enum YwRelativeHorizontalPosition {
        ywRelativeHorizontalPositionMargin = 0,
        ywRelativeHorizontalPositionPage = 1,
        ywRelativeHorizontalPositionColumn = 2,
        ywRelativeHorizontalPositionCharacter = 3,
        ywRelativeHorizontalPositionLeftMarginArea = 4,
        ywRelativeHorizontalPositionRightMarginArea = 5,
        ywRelativeHorizontalPositionInnerMarginArea = 6,
        ywRelativeHorizontalPositionOuterMarginArea = 7

    }
    export const enum YwRelativeVerticalPosition {
        ywRelativeVerticalPositionMargin = 0,
        ywRelativeVerticalPositionPage = 1,
        ywRelativeVerticalPositionParagraph = 2,
        ywRelativeVerticalPositionLine = 3,
        ywRelativeVerticalPositionTopMarginArea = 4,
        ywRelativeVerticalPositionBottomMarginArea = 5,
        ywRelativeVerticalPositionInnerMarginArea = 6,
        ywRelativeVerticalPositionOuterMarginArea = 7

    }
    export const enum YwRelativeVerticalPosition {
        ywRelativeVerticalPositionMargin = 0,
        ywRelativeVerticalPositionPage = 1,
        ywRelativeVerticalPositionParagraph = 2,
        ywRelativeVerticalPositionLine = 3,
        ywRelativeVerticalPositionTopMarginArea = 4,
        ywRelativeVerticalPositionBottomMarginArea = 5,
        ywRelativeVerticalPositionInnerMarginArea = 6,
        ywRelativeVerticalPositionOuterMarginArea = 7

    }
    export const enum YwReplace {
        ywReplaceAll = 2,
        ywReplaceNone = 0,
        ywReplaceOne = 1

    }
    export const enum YwReplace {
        ywReplaceAll = 2,
        ywReplaceNone = 0,
        ywReplaceOne = 1

    }
    export const enum YwRevisedLinesMark {
        ywRevisedLinesMarkNone = 0,
        ywRevisedLinesMarkLeftBorder = 1,
        ywRevisedLinesMarkRightBorder = 2,
        ywRevisedLinesMarkOutsideBorder = 3

    }
    export const enum YwRevisedLinesMark {
        ywRevisedLinesMarkNone = 0,
        ywRevisedLinesMarkLeftBorder = 1,
        ywRevisedLinesMarkRightBorder = 2,
        ywRevisedLinesMarkOutsideBorder = 3

    }
    export const enum YwRevisedPropertiesMark {
        ywRevisedPropertiesMarkNone = 0,
        ywRevisedPropertiesMarkBold = 1,
        ywRevisedPropertiesMarkItalic = 2,
        ywRevisedPropertiesMarkUnderline = 3,
        ywRevisedPropertiesMarkDoubleUnderline = 4,
        ywRevisedPropertiesMarkColorOnly = 5,
        ywRevisedPropertiesMarkStrikeThrough = 6,
        wdRevisedPropertiesMarkDoubleStrikeThrough = 7

    }
    export const enum YwRevisedPropertiesMark {
        ywRevisedPropertiesMarkNone = 0,
        ywRevisedPropertiesMarkBold = 1,
        ywRevisedPropertiesMarkItalic = 2,
        ywRevisedPropertiesMarkUnderline = 3,
        ywRevisedPropertiesMarkDoubleUnderline = 4,
        ywRevisedPropertiesMarkColorOnly = 5,
        ywRevisedPropertiesMarkStrikeThrough = 6,
        wdRevisedPropertiesMarkDoubleStrikeThrough = 7

    }
    export const enum YwRevisionsBalloonPrintOrientation {
        ywBalloonPrintOrientationAuto = 0,
        ywBalloonPrintOrientationPreserve = 1,
        ywBalloonPrintOrientationForceLandscape = 2

    }
    export const enum YwRevisionsBalloonPrintOrientation {
        ywBalloonPrintOrientationAuto = 0,
        ywBalloonPrintOrientationPreserve = 1,
        ywBalloonPrintOrientationForceLandscape = 2

    }
    export const enum YwRevisionsView {
        ywRevisionsViewFinal = 0,
        ywRevisionsViewOriginal = 1

    }
    export const enum YwRevisionsView {
        ywRevisionsViewFinal = 0,
        ywRevisionsViewOriginal = 1

    }
    export const enum YwRevisionType {
        ywNoRevision = 0,
        ywRevisionInsert = 1,
        ywRevisionDelete = 2,
        ywRevisionProperty = 3,
        ywRevisionParagraphNumber = 4,
        ywRevisionDisplayField = 5,
        ywRevisionReconcile = 6,
        ywRevisionConflict = 7,
        ywRevisionStyle = 8,
        ywRevisionReplace = 9,
        ywRevisionParagraphProperty = 10,
        ywRevisionTableProperty = 11,
        ywRevisionSectionProperty = 12,
        ywRevisionStyleDefinition = 13

    }
    export const enum YwRevisionType {
        ywNoRevision = 0,
        ywRevisionInsert = 1,
        ywRevisionDelete = 2,
        ywRevisionProperty = 3,
        ywRevisionParagraphNumber = 4,
        ywRevisionDisplayField = 5,
        ywRevisionReconcile = 6,
        ywRevisionConflict = 7,
        ywRevisionStyle = 8,
        ywRevisionReplace = 9,
        ywRevisionParagraphProperty = 10,
        ywRevisionTableProperty = 11,
        ywRevisionSectionProperty = 12,
        ywRevisionStyleDefinition = 13

    }
    export const enum YwRoutingSlipDelivery {
        ywOneAfterAnother = 0,
        ywAllAtOnce = 1

    }
    export const enum YwRoutingSlipDelivery {
        ywOneAfterAnother = 0,
        ywAllAtOnce = 1

    }
    export const enum YwRoutingSlipStatus {
        ywNotYetRouted = 0,
        ywRouteInProgress = 1,
        ywRouteComplete = 2

    }
    export const enum YwRoutingSlipStatus {
        ywNotYetRouted = 0,
        ywRouteInProgress = 1,
        ywRouteComplete = 2

    }
    export const enum YwRowAlignment {
        ywAlignRowLeft = 0,
        ywAlignRowCenter = 1,
        ywAlignRowRight = 2

    }
    export const enum YwRowAlignment {
        ywAlignRowLeft = 0,
        ywAlignRowCenter = 1,
        ywAlignRowRight = 2

    }
    export const enum YwRowHeightRule {
        ywRowHeightAuto = 0,
        ywRowHeightAtLeast = 1,
        ywRowHeightExactly = 2

    }
    export const enum YwRowHeightRule {
        ywRowHeightAuto = 0,
        ywRowHeightAtLeast = 1,
        ywRowHeightExactly = 2

    }
    export const enum YwRulerStyle {
        ywAdjustNone = 0,
        ywAdjustProportional = 1,
        ywAdjustFirstColumn = 2,
        ywAdjustSameWidth = 3

    }
    export const enum YwRulerStyle {
        ywAdjustNone = 0,
        ywAdjustProportional = 1,
        ywAdjustFirstColumn = 2,
        ywAdjustSameWidth = 3

    }
    export const enum YwSaveFormat {
        ywFormatDocument = 0,
        ywFormatDocument97 = 0,
        ywFormatTemplate = 1,
        ywFormatTemplate97 = 1,
        ywFormatText = 2,
        ywFormatTextLineBreaks = 3,
        ywFormatDOSText = 4,
        ywFormatDOSTextLineBreaks = 5,
        ywFormatRTF = 6,
        ywFormatEncodedText = 7,
        ywFormatUnicodeText = 7,
        ywFormatHTML = 8,
        ywFormatWebArchive = 9,
        ywFormatFilteredHTML = 10,
        ywFormatXML = 11,
        ywFormatXMLDocument = 12,
        ywFormatXMLDocumentMacroEnabled = 13,
        wdFormatXMLTemplate = 14,
        ywFormatXMLTemplateMacroEnabled = 15,
        ywFormatDocumentDefault = 16,
        ywFormatPDF = 17,
        ywFormatXPS = 18

    }
    export const enum YwSaveFormat {
        ywFormatDocument = 0,
        ywFormatDocument97 = 0,
        ywFormatTemplate = 1,
        ywFormatTemplate97 = 1,
        ywFormatText = 2,
        ywFormatTextLineBreaks = 3,
        ywFormatDOSText = 4,
        ywFormatDOSTextLineBreaks = 5,
        ywFormatRTF = 6,
        ywFormatEncodedText = 7,
        ywFormatUnicodeText = 7,
        ywFormatHTML = 8,
        ywFormatWebArchive = 9,
        ywFormatFilteredHTML = 10,
        ywFormatXML = 11,
        ywFormatXMLDocument = 12,
        ywFormatXMLDocumentMacroEnabled = 13,
        wdFormatXMLTemplate = 14,
        ywFormatXMLTemplateMacroEnabled = 15,
        ywFormatDocumentDefault = 16,
        ywFormatPDF = 17,
        ywFormatXPS = 18

    }
    export const enum YwSaveOptions {
        ywDoNotSaveChanges = 0,
        ywPromptToSaveChanges = -2,
        ywSaveChanges = -1

    }
    export const enum YwSaveOptions {
        ywDoNotSaveChanges = 0,
        ywPromptToSaveChanges = -2,
        ywSaveChanges = -1

    }
    export const enum YwScrollbarType {
        ywScrollbarTypeNo = 2,
        ywScrollbarTypeAuto = 0,
        ywScrollbarTypeYes = 1

    }
    export const enum YwScrollbarType {
        ywScrollbarTypeNo = 2,
        ywScrollbarTypeAuto = 0,
        ywScrollbarTypeYes = 1

    }
    export const enum YwSectionDirection {
        ywSectionDirectionLtr = 1,
        ywSectionDirectionRtl = 0

    }
    export const enum YwSectionDirection {
        ywSectionDirectionLtr = 1,
        ywSectionDirectionRtl = 0

    }
    export const enum YwSectionStart {
        ywSectionContinuous = 0,
        ywSectionEvenPage = 3,
        ywSectionNewColumn = 1,
        ywSectionNewPage = 2,
        ywSectionOddPage = 4

    }
    export const enum YwSectionStart {
        ywSectionContinuous = 0,
        ywSectionEvenPage = 3,
        ywSectionNewColumn = 1,
        ywSectionNewPage = 2,
        ywSectionOddPage = 4

    }
    export const enum YwSeekView {
        ywSeekMainDocument = 0,
        ywSeekPrimaryHeader = 1,
        ywSeekFirstPageHeader = 2,
        ywSeekEvenPagesHeader = 3,
        ywSeekPrimaryFooter = 4,
        ywSeekFirstPageFooter = 5,
        ywSeekEvenPagesFooter = 6,
        ywSeekFootnotes = 7,
        ywSeekEndnotes = 8,
        ywSeekCurrentPageHeader = 9,
        ywSeekCurrentPageFooter = 10

    }
    export const enum YwSeekView {
        ywSeekMainDocument = 0,
        ywSeekPrimaryHeader = 1,
        ywSeekFirstPageHeader = 2,
        ywSeekEvenPagesHeader = 3,
        ywSeekPrimaryFooter = 4,
        ywSeekFirstPageFooter = 5,
        ywSeekEvenPagesFooter = 6,
        ywSeekFootnotes = 7,
        ywSeekEndnotes = 8,
        ywSeekCurrentPageHeader = 9,
        ywSeekCurrentPageFooter = 10

    }
    export const enum YwSelectionFlags {
        ywSelStartActive = 1,
        ywSelAtEOL = 2,
        ywSelOvertype = 4,
        ywSelActive = 8,
        ywSelReplace = 16

    }
    export const enum YwSelectionFlags {
        ywSelStartActive = 1,
        ywSelAtEOL = 2,
        ywSelOvertype = 4,
        ywSelActive = 8,
        ywSelReplace = 16

    }
    export const enum YwSelectionType {
        ywNoSelection = 0,
        ywSelectionIP = 1,
        ywSelectionNormal = 2,
        ywSelectionFrame = 3,
        ywSelectionColumn = 4,
        ywSelectionRow = 5,
        ywSelectionBlock = 6,
        ywSelectionInlineShape = 7,
        ywSelectionShape = 8

    }
    export const enum YwSelectionType {
        ywNoSelection = 0,
        ywSelectionIP = 1,
        ywSelectionNormal = 2,
        ywSelectionFrame = 3,
        ywSelectionColumn = 4,
        ywSelectionRow = 5,
        ywSelectionBlock = 6,
        ywSelectionInlineShape = 7,
        ywSelectionShape = 8

    }
    export const enum YwSeparatorType {
        ywSeparatorColon = 2,
        ywSeparatorEmDash = 3,
        ywSeparatorEnDash = 4,
        ywSeparatorHyphen = 0,
        ywSeparatorPeriod = 1

    }
    export const enum YwSeparatorType {
        ywSeparatorColon = 2,
        ywSeparatorEmDash = 3,
        ywSeparatorEnDash = 4,
        ywSeparatorHyphen = 0,
        ywSeparatorPeriod = 1

    }
    export const enum YwSmartTagControlType {
        ywControlActiveX = 13,
        ywControlButton = 6,
        ywControlCheckbox = 9,
        ywControlCombo = 12,
        ywControlDocumentFragment = 14,
        ywControlDocumentFragmentURL = 15,
        ywControlHelp = 3,
        ywControlHelpURL = 4,
        ywControlImage = 8,
        ywControlLabel = 7,
        ywControlLink = 2,
        ywControlListbox = 11,
        ywControlRadioGroup = 16,
        ywControlSeparator = 5,
        ywControlSmartTag = 1,
        ywControlTextbox = 10

    }
    export const enum YwSmartTagControlType {
        ywControlActiveX = 13,
        ywControlButton = 6,
        ywControlCheckbox = 9,
        ywControlCombo = 12,
        ywControlDocumentFragment = 14,
        ywControlDocumentFragmentURL = 15,
        ywControlHelp = 3,
        ywControlHelpURL = 4,
        ywControlImage = 8,
        ywControlLabel = 7,
        ywControlLink = 2,
        ywControlListbox = 11,
        ywControlRadioGroup = 16,
        ywControlSeparator = 5,
        ywControlSmartTag = 1,
        ywControlTextbox = 10

    }
    export const enum YwSortFieldType {
        ywSortFieldSyllable = 3,
        ywSortFieldStroke = 5,
        ywSortFieldNumeric = 1,
        ywSortFieldKoreaKS = 6,
        ywSortFieldJapanJIS = 4,
        ywSortFieldDate = 2,
        ywSortFieldAlphanumeric = 0

    }
    export const enum YwSortFieldType {
        ywSortFieldSyllable = 3,
        ywSortFieldStroke = 5,
        ywSortFieldNumeric = 1,
        ywSortFieldKoreaKS = 6,
        ywSortFieldJapanJIS = 4,
        ywSortFieldDate = 2,
        ywSortFieldAlphanumeric = 0

    }
    export const enum YwSortOrder {
        ywSortOrderAscending = 0,
        ywSortOrderDescending = 1

    }
    export const enum YwSortOrder {
        ywSortOrderAscending = 0,
        ywSortOrderDescending = 1

    }
    export const enum YwSortSeparator {
        ywSortSeparateByCommas = 1,
        ywSortSeparateByDefaultTableSeparator = 2,
        ywSortSeparateByTabs = 0

    }
    export const enum YwSortSeparator {
        ywSortSeparateByCommas = 1,
        ywSortSeparateByDefaultTableSeparator = 2,
        ywSortSeparateByTabs = 0

    }
    export const enum YwSpellingErrorType {
        ywSpellingCapitalization = 2,
        ywSpellingCorrect = 0,
        ywSpellingNotInDictionary = 1

    }
    export const enum YwSpellingErrorType {
        ywSpellingCapitalization = 2,
        ywSpellingCorrect = 0,
        ywSpellingNotInDictionary = 1

    }
    export const enum YwState {
        UNDEFINED = -1,
        FALSE = 0,
        TRUE = 1

    }
    export const enum YwState {
        UNDEFINED = -1,
        FALSE = 0,
        TRUE = 1

    }
    export const enum YwStatistic {
        ywStatisticCharacters = 3,
        ywStatisticCharactersWithSpaces = 5,
        ywStatisticFarEastCharacters = 6,
        ywStatisticLines = 1,
        ywStatisticPages = 2,
        ywStatisticParagraphs = 4,
        ywStatisticWords = 0

    }
    export const enum YwStatistic {
        ywStatisticCharacters = 3,
        ywStatisticCharactersWithSpaces = 5,
        ywStatisticFarEastCharacters = 6,
        ywStatisticLines = 1,
        ywStatisticPages = 2,
        ywStatisticParagraphs = 4,
        ywStatisticWords = 0

    }
    export const enum YwStoryType {
        ywMainTextStory = 1,
        ywFootnotesStory = 2,
        ywEndnotesStory = 3,
        ywCommentsStory = 4,
        ywTextFrameStory = 5,
        ywEvenPagesHeaderStory = 6,
        ywPrimaryHeaderStory = 7,
        ywEvenPagesFooterStory = 8,
        ywPrimaryFooterStory = 9,
        ywFirstPageHeaderStory = 10,
        ywFirstPageFooterStory = 11,
        ywFootnoteSeparatorStory = 12,
        ywFootnoteContinuationSeparatorStory = 13,
        ywFootnoteContinuationNoticeStory = 14,
        ywEndnoteSeparatorStory = 15,
        ywEndnoteContinuationSeparatorStory = 16,
        ywEndnoteContinuationNoticeStory = 17

    }
    export const enum YwStoryType {
        ywMainTextStory = 1,
        ywFootnotesStory = 2,
        ywEndnotesStory = 3,
        ywCommentsStory = 4,
        ywTextFrameStory = 5,
        ywEvenPagesHeaderStory = 6,
        ywPrimaryHeaderStory = 7,
        ywEvenPagesFooterStory = 8,
        ywPrimaryFooterStory = 9,
        ywFirstPageHeaderStory = 10,
        ywFirstPageFooterStory = 11,
        ywFootnoteSeparatorStory = 12,
        ywFootnoteContinuationSeparatorStory = 13,
        ywFootnoteContinuationNoticeStory = 14,
        ywEndnoteSeparatorStory = 15,
        ywEndnoteContinuationSeparatorStory = 16,
        ywEndnoteContinuationNoticeStory = 17

    }
    export const enum YwStyleSheetLinkType {
        ywStyleSheetLinkTypeImported = 1,
        ywStyleSheetLinkTypeLinked = 0

    }
    export const enum YwStyleSheetLinkType {
        ywStyleSheetLinkTypeImported = 1,
        ywStyleSheetLinkTypeLinked = 0

    }
    export const enum YwStyleSheetPrecedence {
        ywStyleSheetPrecedenceHigher = -1,
        ywStyleSheetPrecedenceHighest = 1,
        ywStyleSheetPrecedenceLower = -2,
        ywStyleSheetPrecedenceLowest = 0

    }
    export const enum YwStyleSheetPrecedence {
        ywStyleSheetPrecedenceHigher = -1,
        ywStyleSheetPrecedenceHighest = 1,
        ywStyleSheetPrecedenceLower = -2,
        ywStyleSheetPrecedenceLowest = 0

    }
    export const enum YwStyleType {
        ywStyleTypeCharacter = 2,
        ywStyleTypeList = 4,
        ywStyleTypeParagraph = 1,
        ywStyleTypeTable = 3

    }
    export const enum YwStyleType {
        ywStyleTypeCharacter = 2,
        ywStyleTypeList = 4,
        ywStyleTypeParagraph = 1,
        ywStyleTypeTable = 3

    }
    export const enum YwSummaryMode {
        wdSummaryModeCreateNew = 3,
        wdSummaryModeHideAllButSummary = 1,
        wdSummaryModeHighlight = 0,
        wdSummaryModeInsert = 2

    }
    export const enum YwSummaryMode {
        wdSummaryModeCreateNew = 3,
        wdSummaryModeHideAllButSummary = 1,
        wdSummaryModeHighlight = 0,
        wdSummaryModeInsert = 2

    }
    export const enum YwTabAlignment {
        ywAlignTabLeft = 0,
        ywAlignTabCenter = 1,
        ywAlignTabRight = 2,
        ywAlignTabDecimal = 3,
        ywAlignTabBar = 4,
        ywAlignTabList = 6

    }
    export const enum YwTabAlignment {
        ywAlignTabLeft = 0,
        ywAlignTabCenter = 1,
        ywAlignTabRight = 2,
        ywAlignTabDecimal = 3,
        ywAlignTabBar = 4,
        ywAlignTabList = 6

    }
    export const enum YwTabLeader {
        ywTabLeaderSpaces = 0,
        ywTabLeaderDots = 1,
        ywTabLeaderDashes = 2,
        ywTabLeaderLines = 3,
        ywTabLeaderHeavy = 4,
        ywTabLeaderMiddleDot = 5

    }
    export const enum YwTabLeader {
        ywTabLeaderSpaces = 0,
        ywTabLeaderDots = 1,
        ywTabLeaderDashes = 2,
        ywTabLeaderLines = 3,
        ywTabLeaderHeavy = 4,
        ywTabLeaderMiddleDot = 5

    }
    export const enum YwTableDirection {
        ywTableDirectionRtl = 0,
        ywTableDirectionLtr = 1

    }
    export const enum YwTableDirection {
        ywTableDirectionRtl = 0,
        ywTableDirectionLtr = 1

    }
    export const enum YwTableFieldSeparator {
        ywSeparateByCommas = 2,
        ywSeparateByDefaultListSeparator = 3,
        ywSeparateByParagraphs = 0,
        ywSeparateByTabs = 1

    }
    export const enum YwTableFieldSeparator {
        ywSeparateByCommas = 2,
        ywSeparateByDefaultListSeparator = 3,
        ywSeparateByParagraphs = 0,
        ywSeparateByTabs = 1

    }
    export const enum YwTableFormat {
        ywTableFormatNone = 0,
        ywTableFormatSimple1 = 1,
        ywTableFormatSimple2 = 2,
        ywTableFormatSimple3 = 3,
        ywTableFormatClassic1 = 4,
        ywTableFormatClassic2 = 5,
        ywTableFormatClassic3 = 6,
        ywTableFormatClassic4 = 7,
        ywTableFormatColorful1 = 8,
        ywTableFormatColorful2 = 9,
        ywTableFormatColorful3 = 10,
        ywTableFormatColumns1 = 11,
        ywTableFormatColumns2 = 12,
        ywTableFormatColumns3 = 13,
        ywTableFormatColumns4 = 14,
        ywTableFormatColumns5 = 15,
        ywTableFormatGrid1 = 16,
        ywTableFormatGrid2 = 17,
        ywTableFormatGrid3 = 18,
        ywTableFormatGrid4 = 19,
        ywTableFormatGrid5 = 20,
        ywTableFormatGrid6 = 21,
        ywTableFormatGrid7 = 22,
        ywTableFormatGrid8 = 23,
        ywTableFormatList1 = 24,
        ywTableFormatList2 = 25,
        ywTableFormatList3 = 26,
        ywTableFormatList4 = 27,
        ywTableFormatList5 = 28,
        ywTableFormatList6 = 29,
        ywTableFormatList7 = 30,
        ywTableFormatList8 = 31,
        ywTableFormat3DEffects1 = 32,
        ywTableFormat3DEffects2 = 33,
        ywTableFormat3DEffects3 = 34,
        ywTableFormatContemporary = 35,
        ywTableFormatElegant = 36,
        ywTableFormatProfessional = 37,
        ywTableFormatSubtle1 = 38,
        ywTableFormatSubtle2 = 39,
        ywTableFormatWeb1 = 40,
        ywTableFormatWeb2 = 41,
        ywTableFormatWeb3 = 42

    }
    export const enum YwTableFormat {
        ywTableFormatNone = 0,
        ywTableFormatSimple1 = 1,
        ywTableFormatSimple2 = 2,
        ywTableFormatSimple3 = 3,
        ywTableFormatClassic1 = 4,
        ywTableFormatClassic2 = 5,
        ywTableFormatClassic3 = 6,
        ywTableFormatClassic4 = 7,
        ywTableFormatColorful1 = 8,
        ywTableFormatColorful2 = 9,
        ywTableFormatColorful3 = 10,
        ywTableFormatColumns1 = 11,
        ywTableFormatColumns2 = 12,
        ywTableFormatColumns3 = 13,
        ywTableFormatColumns4 = 14,
        ywTableFormatColumns5 = 15,
        ywTableFormatGrid1 = 16,
        ywTableFormatGrid2 = 17,
        ywTableFormatGrid3 = 18,
        ywTableFormatGrid4 = 19,
        ywTableFormatGrid5 = 20,
        ywTableFormatGrid6 = 21,
        ywTableFormatGrid7 = 22,
        ywTableFormatGrid8 = 23,
        ywTableFormatList1 = 24,
        ywTableFormatList2 = 25,
        ywTableFormatList3 = 26,
        ywTableFormatList4 = 27,
        ywTableFormatList5 = 28,
        ywTableFormatList6 = 29,
        ywTableFormatList7 = 30,
        ywTableFormatList8 = 31,
        ywTableFormat3DEffects1 = 32,
        ywTableFormat3DEffects2 = 33,
        ywTableFormat3DEffects3 = 34,
        ywTableFormatContemporary = 35,
        ywTableFormatElegant = 36,
        ywTableFormatProfessional = 37,
        ywTableFormatSubtle1 = 38,
        ywTableFormatSubtle2 = 39,
        ywTableFormatWeb1 = 40,
        ywTableFormatWeb2 = 41,
        ywTableFormatWeb3 = 42

    }
    export const enum YwTablePosition {
        ywTableBottom = -999997,
        ywTableCenter = -999995,
        ywTableInside = -999994,
        ywTableLeft = -999998,
        ywTableOutside = -999993,
        ywTableRight = -999996,
        ywTableTop = -999999

    }
    export const enum YwTablePosition {
        ywTableBottom = -999997,
        ywTableCenter = -999995,
        ywTableInside = -999994,
        ywTableLeft = -999998,
        ywTableOutside = -999993,
        ywTableRight = -999996,
        ywTableTop = -999999

    }
    export const enum YwTaskPanes {
        ywTaskPaneFormatting = 0,
        ywTaskPaneRevealFormatting = 1,
        ywTaskPaneMailMerge = 2,
        ywTaskPaneTranslate = 3,
        ywTaskPaneSearch = 4,
        ywTaskPaneXMLStructure = 5,
        ywTaskPaneDocumentProtection = 6,
        ywTaskPaneDocumentActions = 7,
        ywTaskPaneSharedWorkspace = 8,
        ywTaskPaneHelp = 9,
        ywTaskPaneResearch = 10,
        ywTaskPaneFaxService = 11,
        ywTaskPaneXMLDocument = 12,
        ywTaskPaneDocumentUpdates = 13,
        ywTaskPaneSignature = 14,
        ywTaskPaneStyleInspector = 15,
        ywTaskPaneApplyStyles = 17

    }
    export const enum YwTaskPanes {
        ywTaskPaneFormatting = 0,
        ywTaskPaneRevealFormatting = 1,
        ywTaskPaneMailMerge = 2,
        ywTaskPaneTranslate = 3,
        ywTaskPaneSearch = 4,
        ywTaskPaneXMLStructure = 5,
        ywTaskPaneDocumentProtection = 6,
        ywTaskPaneDocumentActions = 7,
        ywTaskPaneSharedWorkspace = 8,
        ywTaskPaneHelp = 9,
        ywTaskPaneResearch = 10,
        ywTaskPaneFaxService = 11,
        ywTaskPaneXMLDocument = 12,
        ywTaskPaneDocumentUpdates = 13,
        ywTaskPaneSignature = 14,
        ywTaskPaneStyleInspector = 15,
        ywTaskPaneApplyStyles = 17

    }
    export const enum YwTCSCConverterDirection {
        ywTCSCConverterDirectionAuto = 2,
        ywTCSCConverterDirectionSCTC = 0,
        ywTCSCConverterDirectionTCSC = 1

    }
    export const enum YwTCSCConverterDirection {
        ywTCSCConverterDirectionAuto = 2,
        ywTCSCConverterDirectionSCTC = 0,
        ywTCSCConverterDirectionTCSC = 1

    }
    export const enum YwTextboxTightWrap {
        wdTightAll = 1,
        wdTightFirstAndLastLines = 2,
        wdTightFirstLineOnly = 3,
        wdTightLastLineOnly = 4,
        wdTightNone = 0

    }
    export const enum YwTextboxTightWrap {
        wdTightAll = 1,
        wdTightFirstAndLastLines = 2,
        wdTightFirstLineOnly = 3,
        wdTightLastLineOnly = 4,
        wdTightNone = 0

    }
    export const enum YwTextFormFieldType {
        ywCalculationText = 5,
        ywCurrentDateText = 3,
        ywCurrentTimeText = 4,
        ywDateText = 2,
        ywNumberText = 1,
        ywRegularText = 0

    }
    export const enum YwTextFormFieldType {
        ywCalculationText = 5,
        ywCurrentDateText = 3,
        ywCurrentTimeText = 4,
        ywDateText = 2,
        ywNumberText = 1,
        ywRegularText = 0

    }
    export const enum YwTextOrientation {
        ywTextOrientationHorizontal = 0,
        ywTextOrientationVerticalFarEast = 1,
        ywTextOrientationUpward = 2,
        ywTextOrientationDownward = 3,
        ywTextOrientationHorizontalRotatedFarEast = 4,
        ywTextOrientationVertical = 5

    }
    export const enum YwTextOrientation {
        ywTextOrientationHorizontal = 0,
        ywTextOrientationVerticalFarEast = 1,
        ywTextOrientationUpward = 2,
        ywTextOrientationDownward = 3,
        ywTextOrientationHorizontalRotatedFarEast = 4,
        ywTextOrientationVertical = 5

    }
    export const enum YwTextureIndex {
        ywTexture10Percent = 100,
        ywTexture12Pt5Percent = 125,
        ywTexture15Percent = 150,
        ywTexture17Pt5Percent = 175,
        ywTexture20Percent = 200,
        ywTexture22Pt5Percent = 225,
        ywTexture25Percent = 250,
        ywTexture27Pt5Percent = 275,
        ywTexture2Pt5Percent = 25,
        ywTexture30Percent = 300,
        ywTexture32Pt5Percent = 325,
        ywTexture35Percent = 350,
        ywTexture37Pt5Percent = 375,
        ywTexture40Percent = 400,
        ywTexture42Pt5Percent = 425,
        ywTexture45Percent = 450,
        ywTexture47Pt5Percent = 475,
        ywTexture50Percent = 500,
        ywTexture52Pt5Percent = 525,
        ywTexture55Percent = 550,
        ywTexture57Pt5Percent = 575,
        ywTexture5Percent = 50,
        ywTexture60Percent = 600,
        ywTexture62Pt5Percent = 625,
        ywTexture65Percent = 650,
        ywTexture67Pt5Percent = 675,
        ywTexture70Percent = 700,
        ywTexture72Pt5Percent = 725,
        ywTexture75Percent = 750,
        ywTexture77Pt5Percent = 775,
        ywTexture7Pt5Percent = 75,
        ywTexture80Percent = 800,
        ywTexture82Pt5Percent = 825,
        ywTexture85Percent = 850,
        ywTexture87Pt5Percent = 875,
        ywTexture90Percent = 900,
        ywTexture92Pt5Percent = 925,
        ywTexture95Percent = 950,
        ywTexture97Pt5Percent = 975,
        ywTextureCross = -11,
        ywTextureDarkCross = -5,
        ywTextureDarkDiagonalCross = -6,
        ywTextureDarkDiagonalDown = -3,
        ywTextureDarkDiagonalUp = -4,
        ywTextureDarkHorizontal = -1,
        ywTextureDarkVertical = -2,
        ywTextureDiagonalCross = -12,
        ywTextureDiagonalDown = -9,
        ywTextureDiagonalUp = -10,
        ywTextureHorizontal = -7,
        ywTextureNone = 0,
        ywTextureSolid = 1000,
        ywTextureVertical = -8

    }
    export const enum YwTextureIndex {
        ywTexture10Percent = 100,
        ywTexture12Pt5Percent = 125,
        ywTexture15Percent = 150,
        ywTexture17Pt5Percent = 175,
        ywTexture20Percent = 200,
        ywTexture22Pt5Percent = 225,
        ywTexture25Percent = 250,
        ywTexture27Pt5Percent = 275,
        ywTexture2Pt5Percent = 25,
        ywTexture30Percent = 300,
        ywTexture32Pt5Percent = 325,
        ywTexture35Percent = 350,
        ywTexture37Pt5Percent = 375,
        ywTexture40Percent = 400,
        ywTexture42Pt5Percent = 425,
        ywTexture45Percent = 450,
        ywTexture47Pt5Percent = 475,
        ywTexture50Percent = 500,
        ywTexture52Pt5Percent = 525,
        ywTexture55Percent = 550,
        ywTexture57Pt5Percent = 575,
        ywTexture5Percent = 50,
        ywTexture60Percent = 600,
        ywTexture62Pt5Percent = 625,
        ywTexture65Percent = 650,
        ywTexture67Pt5Percent = 675,
        ywTexture70Percent = 700,
        ywTexture72Pt5Percent = 725,
        ywTexture75Percent = 750,
        ywTexture77Pt5Percent = 775,
        ywTexture7Pt5Percent = 75,
        ywTexture80Percent = 800,
        ywTexture82Pt5Percent = 825,
        ywTexture85Percent = 850,
        ywTexture87Pt5Percent = 875,
        ywTexture90Percent = 900,
        ywTexture92Pt5Percent = 925,
        ywTexture95Percent = 950,
        ywTexture97Pt5Percent = 975,
        ywTextureCross = -11,
        ywTextureDarkCross = -5,
        ywTextureDarkDiagonalCross = -6,
        ywTextureDarkDiagonalDown = -3,
        ywTextureDarkDiagonalUp = -4,
        ywTextureDarkHorizontal = -1,
        ywTextureDarkVertical = -2,
        ywTextureDiagonalCross = -12,
        ywTextureDiagonalDown = -9,
        ywTextureDiagonalUp = -10,
        ywTextureHorizontal = -7,
        ywTextureNone = 0,
        ywTextureSolid = 1000,
        ywTextureVertical = -8

    }
    export const enum YwToaFormat {
        ywTOAClassic = 1,
        ywTOADistinctive = 2,
        ywTOAFormal = 3,
        ywTOASimple = 4,
        ywTOATemplate = 0

    }
    export const enum YwToaFormat {
        ywTOAClassic = 1,
        ywTOADistinctive = 2,
        ywTOAFormal = 3,
        ywTOASimple = 4,
        ywTOATemplate = 0

    }
    export const enum YwTocFormat {
        ywTOCTemplate = 0,
        ywTOCSimple = 6,
        ywTOCModern = 4,
        ywTOCFormal = 5,
        ywTOCFancy = 3,
        ywTOCDistinctive = 2,
        ywTOCClassic = 1

    }
    export const enum YwTocFormat {
        ywTOCTemplate = 0,
        ywTOCSimple = 6,
        ywTOCModern = 4,
        ywTOCFormal = 5,
        ywTOCFancy = 3,
        ywTOCDistinctive = 2,
        ywTOCClassic = 1

    }
    export const enum YwTofFormat {
        ywTOFTemplate = 0,
        ywTOFClassic = 1,
        ywTOFDistinctive = 2,
        ywTOFCentered = 3,
        ywTOFFormal = 4,
        ywTOFSimple = 5

    }
    export const enum YwTofFormat {
        ywTOFTemplate = 0,
        ywTOFClassic = 1,
        ywTOFDistinctive = 2,
        ywTOFCentered = 3,
        ywTOFFormal = 4,
        ywTOFSimple = 5

    }
    export const enum YwTrailingCharacter {
        ywTrailingNone = 2,
        ywTrailingSpace = 1,
        ywTrailingTab = 0

    }
    export const enum YwTrailingCharacter {
        ywTrailingNone = 2,
        ywTrailingSpace = 1,
        ywTrailingTab = 0

    }
    export const enum YwTwoLinesInOneType {
        ywTwoLinesInOneAngleBrackets = 4,
        ywTwoLinesInOneCurlyBrackets = 5,
        ywTwoLinesInOneNoBrackets = 1,
        ywTwoLinesInOneNone = 0,
        ywTwoLinesInOneParentheses = 2,
        ywTwoLinesInOneSquareBrackets = 3

    }
    export const enum YwTwoLinesInOneType {
        ywTwoLinesInOneAngleBrackets = 4,
        ywTwoLinesInOneCurlyBrackets = 5,
        ywTwoLinesInOneNoBrackets = 1,
        ywTwoLinesInOneNone = 0,
        ywTwoLinesInOneParentheses = 2,
        ywTwoLinesInOneSquareBrackets = 3

    }
    export const enum YwUnderline {
        ywUnderlineDash = 7,
        ywUnderlineDashHeavy = 23,
        ywUnderlineDashLong = 39,
        ywUnderlineDashLongHeavy = 55,
        ywUnderlineDotDash = 9,
        ywUnderlineDotDashHeavy = 25,
        ywUnderlineDotDotDash = 10,
        ywUnderlineDotDotDashHeavy = 26,
        ywUnderlineDotted = 4,
        ywUnderlineDottedHeavy = 20,
        ywUnderlineDouble = 3,
        ywUnderlineNone = 0,
        ywUnderlineSingle = 1,
        ywUnderlineThick = 6,
        ywUnderlineWavy = 11,
        ywUnderlineWavyDouble = 43,
        ywUnderlineWavyHeavy = 27,
        ywUnderlineWords = 2

    }
    export const enum YwUnderline {
        ywUnderlineDash = 7,
        ywUnderlineDashHeavy = 23,
        ywUnderlineDashLong = 39,
        ywUnderlineDashLongHeavy = 55,
        ywUnderlineDotDash = 9,
        ywUnderlineDotDashHeavy = 25,
        ywUnderlineDotDotDash = 10,
        ywUnderlineDotDotDashHeavy = 26,
        ywUnderlineDotted = 4,
        ywUnderlineDottedHeavy = 20,
        ywUnderlineDouble = 3,
        ywUnderlineNone = 0,
        ywUnderlineSingle = 1,
        ywUnderlineThick = 6,
        ywUnderlineWavy = 11,
        ywUnderlineWavyDouble = 43,
        ywUnderlineWavyHeavy = 27,
        ywUnderlineWords = 2

    }
    export const enum YwUnits {
        ywCharacter = 1,
        ywWord = 2,
        ywSentence = 3,
        ywParagraph = 4,
        ywLine = 5,
        ywStory = 6,
        ywScreen = 7,
        ywSection = 8,
        ywColumn = 9,
        ywRow = 10,
        ywWindow = 11,
        ywCell = 12,
        ywCharacterFormatting = 13,
        ywParagraphFormatting = 14,
        ywTable = 15,
        ywItem = 16

    }
    export const enum YwUnits {
        ywCharacter = 1,
        ywWord = 2,
        ywSentence = 3,
        ywParagraph = 4,
        ywLine = 5,
        ywStory = 6,
        ywScreen = 7,
        ywSection = 8,
        ywColumn = 9,
        ywRow = 10,
        ywWindow = 11,
        ywCell = 12,
        ywCharacterFormatting = 13,
        ywParagraphFormatting = 14,
        ywTable = 15,
        ywItem = 16

    }
    export const enum YwUseFormattingFrom {
        ywFormattingFromCurrent = 0,
        ywFormattingFromPrompt = 2,
        ywFormattingFromSelected = 1

    }
    export const enum YwUseFormattingFrom {
        ywFormattingFromCurrent = 0,
        ywFormattingFromPrompt = 2,
        ywFormattingFromSelected = 1

    }
    export const enum YwVerticalAlignment {
        ywAlignVerticalBottom = 3,
        ywAlignVerticalCenter = 1,
        ywAlignVerticalJustify = 2,
        ywAlignVerticalTop = 0

    }
    export const enum YwVerticalAlignment {
        ywAlignVerticalBottom = 3,
        ywAlignVerticalCenter = 1,
        ywAlignVerticalJustify = 2,
        ywAlignVerticalTop = 0

    }
    export const enum YwViewType {
        ywMasterView = 5,
        ywNormalView = 1,
        ywOutlineView = 2,
        ywPrintPreview = 4,
        ywPrintView = 3,
        ywReadingView = 7,
        ywWebView = 6

    }
    export const enum YwViewType {
        ywMasterView = 5,
        ywNormalView = 1,
        ywOutlineView = 2,
        ywPrintPreview = 4,
        ywPrintView = 3,
        ywReadingView = 7,
        ywWebView = 6

    }
    export const enum YwVisualSelection {
        ywVisualSelectionBlock = 0,
        ywVisualSelectionContinuous = 1

    }
    export const enum YwVisualSelection {
        ywVisualSelectionBlock = 0,
        ywVisualSelectionContinuous = 1

    }
    export const enum YwWindowState {
        ywWindowStateNormal = 0,
        ywWindowStateMaximize = 1,
        ywWindowStateMinimize = 2

    }
    export const enum YwWindowState {
        ywWindowStateNormal = 0,
        ywWindowStateMaximize = 1,
        ywWindowStateMinimize = 2

    }
    export const enum YwWordDialog {
        HelpWordPerfectHelp = 10,
        GrowFont = 11,
        ShrinkFont = 12,
        Overtype = 13,
        ExtendSelection = 14,
        InsertSpike = 16,
        ChangeCase = 17,
        MoveText = 18,
        CopyText = 19,
        InsertAutoText = 20,
        OtherPane = 21,
        NextWindow = 22,
        PrevWindow = 23,
        RepeatFind = 24,
        NextField = 25,
        PrevField = 26,
        ColumnSelect = 27,
        DeleteWord = 28,
        DeleteBackWord = 29,
        EditClear = 30,
        InsertFieldChars = 31,
        UpdateFields = 32,
        UnlinkFields = 33,
        ToggleFieldDisplay = 34,
        LockFields = 35,
        UnlockFields = 36,
        UpdateSource = 37,
        Indent = 38,
        UnIndent = 39,
        HangingIndent = 40,
        UnHang = 41,
        Font = 42,
        FontSizeSelect = 43,
        WW2_RulerMode = 44,
        Bold = 45,
        Italic = 46,
        SmallCaps = 47,
        AllCaps = 48,
        Strikethrough = 49,
        Hidden = 50,
        Underline = 51,
        DoubleUnderline = 52,
        WordUnderline = 53,
        Superscript = 54,
        Subscript = 55,
        ResetChar = 56,
        CharColor = 57,
        LeftPara = 58,
        CenterPara = 59,
        RightPara = 60,
        JustifyPara = 61,
        SpacePara1 = 62,
        SpacePara15 = 63,
        SpacePara2 = 64,
        CloseUpPara = 65,
        OpenUpPara = 66,
        ResetPara = 67,
        EditRepeat = 68,
        GoBack = 69,
        SaveTemplate = 70,
        OK = 71,
        Cancel = 72,
        CopyFormat = 73,
        PrevPage = 74,
        NextPage = 75,
        NextObject = 76,
        PrevObject = 77,
        DocumentStatistics = 78,
        FileNew = 79,
        FileOpen = 80,
        MailMergeOpenDataSource = 81,
        MailMergeOpenHeaderSource = 82,
        FileSave = 83,
        FileSaveAs = 84,
        FileSaveAll = 85,
        FileSummaryInfo = 86,
        FileTemplates = 87,
        FilePrint = 88,
        FilePrintPreview = 89,
        WW2_PrintMergeHelper = 95,
        FilePrintSetup = 97,
        FileExit = 98,
        FileFind = 99,
        FormatAddrFonts = 103,
        WW2_PrintMergeCreateDataSource = 105,
        WW2_PrintMergeCreateHeaderSource = 106,
        EditUndo = 107,
        EditPaste = 110,
        EditPasteSpecial = 111,
        EditFind = 112,
        EditFindFont = 113,
        EditFindPara = 114,
        EditFindStyle = 115,
        EditFindClearFormatting = 116,
        EditReplace = 117,
        EditReplaceFont = 118,
        EditReplacePara = 119,
        EditReplaceStyle = 120,
        EditReplaceClearFormatting = 121,
        WW7_EditGoTo = 122,
        WW7_EditAutoText = 123,
        TableInsertTable = 129,
        ViewFieldCodes = 150,
        Style = 151,
        ToolsCustomize = 152,
        ViewRuler = 153,
        ViewStatusBar = 154,
        NormalViewHeaderArea = 155,
        ViewFootnoteArea = 156,
        ViewAnnotations = 157,
        InsertFrame = 158,
        InsertBreak = 159,
        WW2_InsertFootnote = 160,
        InsertAnnotation = 161,
        InsertSymbol = 162,
        InsertPicture = 163,
        InsertFile = 164,
        InsertDateTime = 165,
        InsertField = 166,
        EditBookmark = 168,
        MarkIndexEntry = 169,
        InsertIndex = 170,
        InsertTableOfContents = 171,
        InsertObject = 172,
        ToolsCreateEnvelope = 173,
        FormatFont = 174,
        FormatParagraph = 175,
        FormatSectionLayout = 176,
        FormatColumns = 177,
        FilePageSetup = 178,
        FormatTabs = 179,
        FormatStyle = 180,
        FormatDefineStyleFont = 181,
        FormatDefineStylePara = 182,
        FormatDefineStyleTabs = 183,
        FormatDefineStyleFrame = 184,
        FormatDefineStyleBorders = 185,
        FormatDefineStyleLang = 186,
        ToolsLanguage = 188,
        FormatBordersAndShading = 189,
        ToolsSpelling = 191,
        ToolsSpellSelection = 192,
        ToolsGrammar = 193,
        ToolsThesaurus = 194,
        ToolsHyphenation = 195,
        ToolsBulletsNumbers = 196,
        ToolsRevisions = 197,
        ToolsCompareVersions = 198,
        TableSort = 199,
        ToolsRepaginate = 201,
        WW7_ToolsOptions = 202,
        ToolsOptionsGeneral = 203,
        ToolsOptionsView = 204,
        ToolsAdvancedSettings = 206,
        ToolsOptionsPrint = 208,
        ToolsOptionsSave = 209,
        WW2_ToolsOptionsToolbar = 210,
        ToolsOptionsSpelling = 211,
        ToolsOptionsGrammar = 212,
        ToolsOptionsUserInfo = 213,
        ToolsMacro = 215,
        WindowNewWindow = 217,
        WindowArrangeAll = 218,
        WindowList = 220,
        FormatRetAddrFonts = 221,
        Organizer = 222,
        ToolsOptionsEdit = 224,
        ToolsOptionsFileLocations = 225,
        ToolsAutoCorrectSmartQuotes = 227,
        ToolsWordCount = 228,
        DocSplit = 229,
        DocSize = 230,
        InsertPageNumbers = 294,
        COUNT = 231

    }
    export const enum YwWordDialog {
        HelpWordPerfectHelp = 10,
        GrowFont = 11,
        ShrinkFont = 12,
        Overtype = 13,
        ExtendSelection = 14,
        InsertSpike = 16,
        ChangeCase = 17,
        MoveText = 18,
        CopyText = 19,
        InsertAutoText = 20,
        OtherPane = 21,
        NextWindow = 22,
        PrevWindow = 23,
        RepeatFind = 24,
        NextField = 25,
        PrevField = 26,
        ColumnSelect = 27,
        DeleteWord = 28,
        DeleteBackWord = 29,
        EditClear = 30,
        InsertFieldChars = 31,
        UpdateFields = 32,
        UnlinkFields = 33,
        ToggleFieldDisplay = 34,
        LockFields = 35,
        UnlockFields = 36,
        UpdateSource = 37,
        Indent = 38,
        UnIndent = 39,
        HangingIndent = 40,
        UnHang = 41,
        Font = 42,
        FontSizeSelect = 43,
        WW2_RulerMode = 44,
        Bold = 45,
        Italic = 46,
        SmallCaps = 47,
        AllCaps = 48,
        Strikethrough = 49,
        Hidden = 50,
        Underline = 51,
        DoubleUnderline = 52,
        WordUnderline = 53,
        Superscript = 54,
        Subscript = 55,
        ResetChar = 56,
        CharColor = 57,
        LeftPara = 58,
        CenterPara = 59,
        RightPara = 60,
        JustifyPara = 61,
        SpacePara1 = 62,
        SpacePara15 = 63,
        SpacePara2 = 64,
        CloseUpPara = 65,
        OpenUpPara = 66,
        ResetPara = 67,
        EditRepeat = 68,
        GoBack = 69,
        SaveTemplate = 70,
        OK = 71,
        Cancel = 72,
        CopyFormat = 73,
        PrevPage = 74,
        NextPage = 75,
        NextObject = 76,
        PrevObject = 77,
        DocumentStatistics = 78,
        FileNew = 79,
        FileOpen = 80,
        MailMergeOpenDataSource = 81,
        MailMergeOpenHeaderSource = 82,
        FileSave = 83,
        FileSaveAs = 84,
        FileSaveAll = 85,
        FileSummaryInfo = 86,
        FileTemplates = 87,
        FilePrint = 88,
        FilePrintPreview = 89,
        WW2_PrintMergeHelper = 95,
        FilePrintSetup = 97,
        FileExit = 98,
        FileFind = 99,
        FormatAddrFonts = 103,
        WW2_PrintMergeCreateDataSource = 105,
        WW2_PrintMergeCreateHeaderSource = 106,
        EditUndo = 107,
        EditPaste = 110,
        EditPasteSpecial = 111,
        EditFind = 112,
        EditFindFont = 113,
        EditFindPara = 114,
        EditFindStyle = 115,
        EditFindClearFormatting = 116,
        EditReplace = 117,
        EditReplaceFont = 118,
        EditReplacePara = 119,
        EditReplaceStyle = 120,
        EditReplaceClearFormatting = 121,
        WW7_EditGoTo = 122,
        WW7_EditAutoText = 123,
        TableInsertTable = 129,
        ViewFieldCodes = 150,
        Style = 151,
        ToolsCustomize = 152,
        ViewRuler = 153,
        ViewStatusBar = 154,
        NormalViewHeaderArea = 155,
        ViewFootnoteArea = 156,
        ViewAnnotations = 157,
        InsertFrame = 158,
        InsertBreak = 159,
        WW2_InsertFootnote = 160,
        InsertAnnotation = 161,
        InsertSymbol = 162,
        InsertPicture = 163,
        InsertFile = 164,
        InsertDateTime = 165,
        InsertField = 166,
        EditBookmark = 168,
        MarkIndexEntry = 169,
        InsertIndex = 170,
        InsertTableOfContents = 171,
        InsertObject = 172,
        ToolsCreateEnvelope = 173,
        FormatFont = 174,
        FormatParagraph = 175,
        FormatSectionLayout = 176,
        FormatColumns = 177,
        FilePageSetup = 178,
        FormatTabs = 179,
        FormatStyle = 180,
        FormatDefineStyleFont = 181,
        FormatDefineStylePara = 182,
        FormatDefineStyleTabs = 183,
        FormatDefineStyleFrame = 184,
        FormatDefineStyleBorders = 185,
        FormatDefineStyleLang = 186,
        ToolsLanguage = 188,
        FormatBordersAndShading = 189,
        ToolsSpelling = 191,
        ToolsSpellSelection = 192,
        ToolsGrammar = 193,
        ToolsThesaurus = 194,
        ToolsHyphenation = 195,
        ToolsBulletsNumbers = 196,
        ToolsRevisions = 197,
        ToolsCompareVersions = 198,
        TableSort = 199,
        ToolsRepaginate = 201,
        WW7_ToolsOptions = 202,
        ToolsOptionsGeneral = 203,
        ToolsOptionsView = 204,
        ToolsAdvancedSettings = 206,
        ToolsOptionsPrint = 208,
        ToolsOptionsSave = 209,
        WW2_ToolsOptionsToolbar = 210,
        ToolsOptionsSpelling = 211,
        ToolsOptionsGrammar = 212,
        ToolsOptionsUserInfo = 213,
        ToolsMacro = 215,
        WindowNewWindow = 217,
        WindowArrangeAll = 218,
        WindowList = 220,
        FormatRetAddrFonts = 221,
        Organizer = 222,
        ToolsOptionsEdit = 224,
        ToolsOptionsFileLocations = 225,
        ToolsAutoCorrectSmartQuotes = 227,
        ToolsWordCount = 228,
        DocSplit = 229,
        DocSize = 230,
        InsertPageNumbers = 294,
        COUNT = 231

    }
    export const enum YwWrapSideType {
        ywWrapBoth = 0,
        ywWrapLargest = 3,
        ywWrapLeft = 1,
        ywWrapRight = 2

    }
    export const enum YwWrapSideType {
        ywWrapBoth = 0,
        ywWrapLargest = 3,
        ywWrapLeft = 1,
        ywWrapRight = 2

    }
    export const enum YwWrapType {
        ywWrapInline = 7,
        ywWrapNone = 3,
        ywWrapSquare = 0,
        ywWrapThrough = 2,
        ywWrapTight = 1,
        ywWrapTopBottom = 4,
        ywWrapBehind = 5,
        ywWrapFront = 3

    }
    export const enum YwWrapType {
        ywWrapInline = 7,
        ywWrapNone = 3,
        ywWrapSquare = 0,
        ywWrapThrough = 2,
        ywWrapTight = 1,
        ywWrapTopBottom = 4,
        ywWrapBehind = 5,
        ywWrapFront = 3

    }
    export const enum YwWrapTypeMerged {
        ywWrapMergeInline = 0,
        ywWrapMergeSquare = 1,
        ywWrapMergeTight = 2,
        ywWrapMergeBehind = 3,
        ywWrapMergeFront = 4,
        ywWrapMergeThrough = 5,
        ywWrapMergeTopBottom = 6

    }
    export const enum YwWrapTypeMerged {
        ywWrapMergeInline = 0,
        ywWrapMergeSquare = 1,
        ywWrapMergeTight = 2,
        ywWrapMergeBehind = 3,
        ywWrapMergeFront = 4,
        ywWrapMergeThrough = 5,
        ywWrapMergeTopBottom = 6

    }
}
declare namespace word.dialog {
    export interface FindReplaceDialog {
        setFind(findStr: string): void
        setReplace(replaceStr: string): void
    }
    export interface FindReplaceDialog {
        setFind(findStr: string): void
        setReplace(replaceStr: string): void
    }
}
declare namespace word.event {
    export interface ApplicationAdapter {
        quit(e: word.event.ApplicationEvent): void
        windowActivate(e: word.event.ApplicationEvent): void
        windowDeactivate(e: word.event.ApplicationEvent): void
        documentChange(e: word.event.ApplicationEvent): void
        documentOpen(e: word.event.ApplicationEvent): void
        windowSize(e: word.event.ApplicationEvent): void
        newDocument(e: word.event.ApplicationEvent): void
        documentBeforePrint(e: word.event.ApplicationEvent): void
        documentBeforeSave(e: word.event.ApplicationEvent): void
        documentBeforeClose(e: word.event.ApplicationEvent): void
        windowSelectionChange(e: word.event.ApplicationEvent): void
        mailMergeAfterRecordMerge(e: word.event.ApplicationEvent): void
        mailMergeBeforeRecordMerge(e: word.event.ApplicationEvent): void
        mailMergeDataSourceValidate(e: word.event.ApplicationEvent): void
        mailMergeWizardSendToCustom(e: word.event.ApplicationEvent): void
        mailMergeWizardStateChange(e: word.event.ApplicationEvent): void
        ePostagePropertyDialog(e: word.event.ApplicationEvent): void
        mailMergeAfterMerge(e: word.event.ApplicationEvent): void
        mailMergeBeforeMerge(e: word.event.ApplicationEvent): void
        mailMergeDataSourceLoad(e: word.event.ApplicationEvent): void
        windowBeforeDoubleClick(e: word.event.ApplicationEvent): void
        windowBeforeRightClick(e: word.event.ApplicationEvent): void
        xMLSelectionChange(e: word.event.ApplicationEvent): void
        xMLValidationError(e: word.event.ApplicationEvent): void
        documentSync(e: word.event.ApplicationEvent): void
        ePostageInsert(e: word.event.ApplicationEvent): void
        ePostageInsertEx(e: word.event.ApplicationEvent): void
    }
    export interface ApplicationAdapter {
        quit(e: word.event.ApplicationEvent): void
        windowActivate(e: word.event.ApplicationEvent): void
        windowDeactivate(e: word.event.ApplicationEvent): void
        documentChange(e: word.event.ApplicationEvent): void
        documentOpen(e: word.event.ApplicationEvent): void
        windowSize(e: word.event.ApplicationEvent): void
        newDocument(e: word.event.ApplicationEvent): void
        documentBeforePrint(e: word.event.ApplicationEvent): void
        documentBeforeSave(e: word.event.ApplicationEvent): void
        documentBeforeClose(e: word.event.ApplicationEvent): void
        windowSelectionChange(e: word.event.ApplicationEvent): void
        mailMergeAfterRecordMerge(e: word.event.ApplicationEvent): void
        mailMergeBeforeRecordMerge(e: word.event.ApplicationEvent): void
        mailMergeDataSourceValidate(e: word.event.ApplicationEvent): void
        mailMergeWizardSendToCustom(e: word.event.ApplicationEvent): void
        mailMergeWizardStateChange(e: word.event.ApplicationEvent): void
        ePostagePropertyDialog(e: word.event.ApplicationEvent): void
        mailMergeAfterMerge(e: word.event.ApplicationEvent): void
        mailMergeBeforeMerge(e: word.event.ApplicationEvent): void
        mailMergeDataSourceLoad(e: word.event.ApplicationEvent): void
        windowBeforeDoubleClick(e: word.event.ApplicationEvent): void
        windowBeforeRightClick(e: word.event.ApplicationEvent): void
        xMLSelectionChange(e: word.event.ApplicationEvent): void
        xMLValidationError(e: word.event.ApplicationEvent): void
        documentSync(e: word.event.ApplicationEvent): void
        ePostageInsert(e: word.event.ApplicationEvent): void
        ePostageInsertEx(e: word.event.ApplicationEvent): void
    }
    export interface ApplicationEvent {
        getType(): int
        getSelection(): word.Selection
        getDocument(): word.Document
        isCancel(): boolean
        setCancel(cancel: boolean): void
        setShowSaveAsUI(show: boolean): void
        isShowSaveAsUI(): boolean
        getWindow(): word.Window
    }
    export interface ApplicationEvent {
        getType(): int
        getSelection(): word.Selection
        getDocument(): word.Document
        isCancel(): boolean
        setCancel(cancel: boolean): void
        setShowSaveAsUI(show: boolean): void
        isShowSaveAsUI(): boolean
        getWindow(): word.Window
    }
    export interface ApplicationListener {
        quit(e: word.event.ApplicationEvent): void
        windowActivate(e: word.event.ApplicationEvent): void
        windowDeactivate(e: word.event.ApplicationEvent): void
        documentChange(e: word.event.ApplicationEvent): void
        documentOpen(e: word.event.ApplicationEvent): void
        windowSize(e: word.event.ApplicationEvent): void
        newDocument(e: word.event.ApplicationEvent): void
        documentBeforePrint(e: word.event.ApplicationEvent): void
        documentBeforeSave(e: word.event.ApplicationEvent): void
        documentBeforeClose(e: word.event.ApplicationEvent): void
        windowSelectionChange(e: word.event.ApplicationEvent): void
        mailMergeAfterRecordMerge(e: word.event.ApplicationEvent): void
        mailMergeBeforeRecordMerge(e: word.event.ApplicationEvent): void
        mailMergeDataSourceValidate(e: word.event.ApplicationEvent): void
        mailMergeWizardSendToCustom(e: word.event.ApplicationEvent): void
        mailMergeWizardStateChange(e: word.event.ApplicationEvent): void
        ePostagePropertyDialog(e: word.event.ApplicationEvent): void
        mailMergeAfterMerge(e: word.event.ApplicationEvent): void
        mailMergeBeforeMerge(e: word.event.ApplicationEvent): void
        mailMergeDataSourceLoad(e: word.event.ApplicationEvent): void
        windowBeforeDoubleClick(e: word.event.ApplicationEvent): void
        windowBeforeRightClick(e: word.event.ApplicationEvent): void
        xMLSelectionChange(e: word.event.ApplicationEvent): void
        xMLValidationError(e: word.event.ApplicationEvent): void
        documentSync(e: word.event.ApplicationEvent): void
        ePostageInsert(e: word.event.ApplicationEvent): void
        ePostageInsertEx(e: word.event.ApplicationEvent): void
    }
    export interface ApplicationListener {
        quit(e: word.event.ApplicationEvent): void
        windowActivate(e: word.event.ApplicationEvent): void
        windowDeactivate(e: word.event.ApplicationEvent): void
        documentChange(e: word.event.ApplicationEvent): void
        documentOpen(e: word.event.ApplicationEvent): void
        windowSize(e: word.event.ApplicationEvent): void
        newDocument(e: word.event.ApplicationEvent): void
        documentBeforePrint(e: word.event.ApplicationEvent): void
        documentBeforeSave(e: word.event.ApplicationEvent): void
        documentBeforeClose(e: word.event.ApplicationEvent): void
        windowSelectionChange(e: word.event.ApplicationEvent): void
        mailMergeAfterRecordMerge(e: word.event.ApplicationEvent): void
        mailMergeBeforeRecordMerge(e: word.event.ApplicationEvent): void
        mailMergeDataSourceValidate(e: word.event.ApplicationEvent): void
        mailMergeWizardSendToCustom(e: word.event.ApplicationEvent): void
        mailMergeWizardStateChange(e: word.event.ApplicationEvent): void
        ePostagePropertyDialog(e: word.event.ApplicationEvent): void
        mailMergeAfterMerge(e: word.event.ApplicationEvent): void
        mailMergeBeforeMerge(e: word.event.ApplicationEvent): void
        mailMergeDataSourceLoad(e: word.event.ApplicationEvent): void
        windowBeforeDoubleClick(e: word.event.ApplicationEvent): void
        windowBeforeRightClick(e: word.event.ApplicationEvent): void
        xMLSelectionChange(e: word.event.ApplicationEvent): void
        xMLValidationError(e: word.event.ApplicationEvent): void
        documentSync(e: word.event.ApplicationEvent): void
        ePostageInsert(e: word.event.ApplicationEvent): void
        ePostageInsertEx(e: word.event.ApplicationEvent): void
    }
    export interface DocumentEvent {
        getDocument(): word.Document
    }
    export interface DocumentEvent {
        getDocument(): word.Document
    }
    export interface DocumentListener {
        close(e: word.event.DocumentEvent): void
        open(e: word.event.DocumentEvent): void
    }
    export interface DocumentListener {
        close(e: word.event.DocumentEvent): void
        open(e: word.event.DocumentEvent): void
    }
    export interface MainAdapter {
        componentResized(e: any): void
        componentMoved(e: any): void
        componentShown(e: any): void
        componentHidden(e: any): void
        mousePressed(e: any): void
        mouseReleased(e: any): void
        mouseClicked(e: any): void
        mouseExited(e: any): void
        mouseEntered(e: any): void
        windowOpened(e: any): void
        windowClosing(e: any): void
        windowClosed(e: any): void
        windowIconified(e: any): void
        windowDeiconified(e: any): void
        windowActivated(e: any): void
        windowDeactivated(e: any): void
        statusChanged(event: any): boolean
        windowActivate(event: any): void
        beforeExit(event: any): boolean
        saveTree(event: any): void
        loadTree(event: any): void
        windowDeactivate(event: any): void
        windowResize(event: any): void
        afterOpenBinder(event: any): void
        workbookActivate(event: any): void
        saveTemplate(event: any): void
        loadTemplate(event: any): void
        delTemplate(event: any): void
        loadCustomMeta(event: any): void
        loadFixedBase(event: any): void
        wordStatusChange(event: any): boolean
        caretUpdate(e: any): void
        beforeCloseBinder(event: any): boolean
        afterCreateBinder(event: any): void
        beforePrintWorkbook(event: any): boolean
        afterPrintWorkbook(event: any): boolean
        beforeSaveWorkbook(event: any): boolean
        workbookDeactivate(event: any): void
        afterBinderReveredSave(event: any): void
    }
    export interface MainAdapter {
        componentResized(e: any): void
        componentMoved(e: any): void
        componentShown(e: any): void
        componentHidden(e: any): void
        mousePressed(e: any): void
        mouseReleased(e: any): void
        mouseClicked(e: any): void
        mouseExited(e: any): void
        mouseEntered(e: any): void
        windowOpened(e: any): void
        windowClosing(e: any): void
        windowClosed(e: any): void
        windowIconified(e: any): void
        windowDeiconified(e: any): void
        windowActivated(e: any): void
        windowDeactivated(e: any): void
        statusChanged(event: any): boolean
        windowActivate(event: any): void
        beforeExit(event: any): boolean
        saveTree(event: any): void
        loadTree(event: any): void
        windowDeactivate(event: any): void
        windowResize(event: any): void
        afterOpenBinder(event: any): void
        workbookActivate(event: any): void
        saveTemplate(event: any): void
        loadTemplate(event: any): void
        delTemplate(event: any): void
        loadCustomMeta(event: any): void
        loadFixedBase(event: any): void
        wordStatusChange(event: any): boolean
        caretUpdate(e: any): void
        beforeCloseBinder(event: any): boolean
        afterCreateBinder(event: any): void
        beforePrintWorkbook(event: any): boolean
        afterPrintWorkbook(event: any): boolean
        beforeSaveWorkbook(event: any): boolean
        workbookDeactivate(event: any): void
        afterBinderReveredSave(event: any): void
    }
}
declare namespace word.resource {
    export const enum DocumentException {
        //INSERT_STRING_FAILURE=�����ַ�ʧ��!,
        //DELETE_STRING_FAILURE=ɾ���ַ�ʧ��!,
        //OFFSET_ERROR=��ʼβע����Ϊ����!,
        //DOCUMENT_NO_FOOTNOTE=���½�û�н�ע,
        //FOOTNOTE_ERROR1=����Խ�硣���½ڹ��� [0,,
        //FOOTNOTE_ERROR2=] ����ע��û���ҵ�����ָ���ĵ� ,
        //FOOTNOTE_ERROR3= ����ע��,
        //DOCUMENT_NO_ENDNOTE=���½�û��βע,
        //ENDNOTE_ERROR1=����Խ�硣���½ڹ��� [0,,
        //ENDNOTE_ERROR2=] ��βע��û���ҵ�����ָ���ĵ� ,
        //ENDNOTE_ERROR3= ��βע��,
        //OFFSET_IS_MINUS=�ַ�λ�ò���Ϊ����,
        //OFFSET_IS_OVER=�ַ�λ���Ѿ��������ĵ����ȡ�(�ĵ����� = ,
        //TABLEE_ERROR=û������ָ���ı���,
        //SSTABLEE_ERROR=û������ָ���Ĺ�����,
        //SHURU_LENGTH=������Ĺ��λ��,
        //DAYU_DOCLENGTH=���ڸ��½ڵ��ı�����,
        //INSERT_FIELD_TYPE_RANGE=���������� [0,,
        //INSERT_FIELD_NAME_RANGE=������и����������� [0,,
        //INSERT_FIELD_FORMAT_RANGE=����������������ʽ������ [0,,
        //BOOKNAME_ERROR1=��ǩ���ֲ���Ϊ�ջ�ո�,
        //BOOKNAME_ERROR2=��ǩ���ֿ�ͷ��һ����ĸ����Ϊ���� ''0-9'' �� ''_'',
        //BOOKNAME_ERROR3=��ǩ�����к��зǷ��ַ�,
        //TABSTOPS_ERROR1=����Խ�硣�ö����Ʊ�λ�� [0,,
        //TABSTOPS_ERROR2=] ��,������Ϸ����Ʊ�λ�����š�,
        //SECTION_ERROR1=����Խ�硣ȫ�Ĺ��� [0,,
        //SECTION_ERROR2=] ���½ڣ�û���ҵ�����ָ���ĵ� ,
        //SECTION_ERROR3= ���½�,
        //FILENAME_ERROR1=�ļ�������Ϊ�ա�,
        //FOOTNOTE_CANNOT_INSERT_OBJECT=��עβע���ܲ���ͼƬ�������ļ���Ӱ���ļ���,
        //FOOTNOTE_NO_AUTOFORMAT=��עβעû���Զ���ʽ���Ĺ���,
        //CANNOTPASTELINK=������ճ�����ӵ�����,
        //NOCOMPONENT=û����ָ�������,
        //NODATETIME=û����ָ����ʱ�����ڸ�ʽ,
        //NO_PHONETIC=����״̬����ȡƴ��,
        //HAS_STYLENAME=�Ѿ����ڣ����ѱ�����Ϊ������ʽ.,
        //NOHAS_STYLE=�����ڣ����߲����û��Զ�����ʽ������ɾ��,
        //NOHAS_STYLE1=��ʽ�����ڡ�ȡĬ����ʽ:,
        //HYPERLINK_NO_BOOKMARK=����ָ�����ĵ�û����ǩ,
        //HYPERLINK_NO_HEADING=����ָ�����ĵ�û�б���,
        //PROTECED_DOC=���ĵ����ڱ���״̬������ȡ���ĵ�����,
        //PROTECED_TRACKCHANGE=���ĵ����ڱ����޶�״̬������ȡ���ĵ�����,
        //UNPROTECED_DOC=���ĵ����ڲ�����״̬���ĵ��Ѿ��Ǳ��Ᵽ����,
        //NO_MERGE=ѡ�������Ǿ������򣬲��ܺϲ���Ԫ��������ѡ��,
        //NO_SPLIT=ѡ�������кϲ���Ԫ���Ҳ��Ǿ������򣬲��ܲ�ֵ�Ԫ��������ѡ��,
        //NOIN_TABLE=�����ڱ������ٴβ������������ѡ��������λ��,
        //NO_SHADING=���ܶ�ĳ���½ڻ�����ƪ�ĵ����õ���,
        //NOT_COLLAPSE_OR_EXPAND=�����ɱ����ڣ����ܶԶ�������۵���չ�������ı������,
        //NOT_MOVEUP_OR_DOWN=�����ɱ����ڣ����ܶԶ������ƻ�����,
        //NO_HEADANDFOOT=�����ĵ������ҳü��ҳ�Ź�����Ч,
        //SAME_NAME=�ñ༭�������Ѿ�������ͬ���ֵı���,
        //NO_TABLE_AS_NAME=û��ָ�����Ƶı���,
        //POSITION_OUTOF_TABLE=ָ����λ���ڱ����в�����,
        //UNFIT_ACTION_AREA=�����в���ִ�еĶ���ָ��,
        //VALUE_OUT_OF_BOUNDS=���������ѳ�����Χ,
        //DOC_EXIST=�ĵ��Ѿ�����,
        //NAME_INVALID=�ĵ������к��зǷ��ַ�,
        //NAME_TOO_LONG=�ĵ����ƹ�������಻���� 31 ���ַ�,
        //ERROR_FIELD_CODE=���ź�����벻����,
        //WP_NOT_OPEN=��ǰ���ִ���Ӧ�ó���δ������,
        //BARCODE_ERROR1=�����������Ϣ̫��,
        //BARCODE_ERROR2=����������������,
        //ERROR_MESSAGE3=�˷���ֻ��Ӧ�������Ĳ���,
        //ERROR_MESSAGE4=�˷���ֻ��Ӧ����ҳü��ҳ�Ż����Ĳ���,
        //FOOTENDNOT_ERROR1=��עβע���Զ����Ǳ���Ϊһ�ַ�ֵ,
        //FOOTENDNOTE_ERROR2=���ý�עβע�ı�Ÿ�ʽǰ����ʹ���Զ�����Ϊ��,
        //WBL_BULLET_ONLY=�˷���ֻ��Ӧ��������Ϊ��Ŀ���ŵ��б�����,
        //WBL_NUMBER_OUTLINENUMBER=�˷���ֻ��Ӧ��������Ϊ��Ż�༶��ŵ��б�����,
        //WBL_OUTLINENUMBER_ONLY=�˷���ֻ��Ӧ��������Ϊ�༶��ŵ��б�����,
        //SPREADSHEET_ONLY=�˷���ֻ��Ӧ���ڵ��ӱ�����,
        //WORDPROCESSOR_ONLY=�˷���ֻ��Ӧ�������ִ�����,
        //PRESENTATION_ONLY=�˷���ֻ��Ӧ���ڼ�������,
        //LETTER_STYLE_ONLY=ֻ�е���ֽ�ż�Ϊ��ֽ��ʽʱ�˷�������Ч,
        //GRID_LETTER_STYLE=�˷���ֻ��Ӧ���ڸ�ֽ�ż�Ϊ�������ֽ��ʽ,
        //COLUMN_WIDTH_NOTSET=��ǰ����Ϊ���������ʱ�������÷����е�������,
        //COLUMN_SPACE_NOTSET=��ǰ����Ϊ���������ʱ�������÷����е������,
        //GRID_SETTING_NONGRID=�˷�������Ӧ�����ĵ���������Ϊ��,
        //GRID_SETTING_LINESPACING1=�˷���ֻ��Ӧ���ڶ����о���������Сֵ���̶�ֵ��౶�о�,
        //PROTECTED_SECTION=�ýڴ��ڱ���״̬������ȡ������,
        //CAN_NOT_EDIT=ָ��λ�ô��ڲ��ɱ༭״̬,
        //FREETABLE_INSERT_HYPERLINK_ERROR=�����ɱ�����,�޷�ͬʱ�ڶ����Ԫ���в��볬����,
        //DOCUMENT_COUNT_ERROR=���ֻ�ܲ��� 64 ���ĵ�,
        //GRID_SETTING_NO_PAGEMARGE=ֻ���ڲ�ʹ��ҳ�߾�״̬�´˷�������Ч,
        //GRID_SETTING_HORIZONTAL=ֻ������ʾˮƽ�����״̬�´˷�������Ч,
        //GRID_SETTING_VERTICAL=ֻ������ʾ��ֱ�����״̬�´˷�������Ч,
        //CELL_IS_OVER=�ַ�λ���Ѿ������õ�Ԫ���ַ����ȡ�(��Ԫ���ַ����� = ,
        //WATERMARK_PICTURE=�˷���ֻ����ˮӡ����ΪͼƬˮӡʱ����Ч,
        //WATERMARK_TEXT=�˷���ֻ����ˮӡ����Ϊ����ˮӡʱ����Ч,
        //NO_BASESTYLE=��׼��ʽ������,
        //VALIDATE_DATE=��ǰ���ݱ���Ϊ��Ч����,
        //LESS_THAN_40=���ֶ������ֻ������ 40 ���ֽ�,
        //LESS_THAN_110=���ֶ������ֻ������ 110 ���ֽ�,
        //LESS_THAN_20=���ֶ������ֻ������ 20 ���ֽ�,
        //MORE_THAN_8=���ֶ����������� 8 ���ֽ�,
        //BARCODE_NUMBER=�������š�Ϊ��ѡ�ֶΣ�����������,
        //BARCODE_PRODUCE=������������λ��Ϊ��ѡ�ֶΣ�����������,
        //CROSSREF_ERROR=�������ô���,
        //LESS_THAN_30=���ֶ����ֻ������ 30 ���ֽ�,
        //LESS_THAN_150=���ֶ����ֻ������ 150 ���ֽ�,
        //FILE_FORMAT_INVALID=���ļ��ĸ�ʽ����ʶ��,
        //PAPER_DIMENSION_INVALID=ֽ�ųߴ�����������Ч,
        //FILE_DIRECTORY_NOT_EXIST=�ļ�Ŀ¼������,
        //LESS_THAN_8=���ֶβ��ܳ��� 8 ���ֽ�,
        //BOOKMARK_NOT_EXIST=�����ڴ���ǩ��,
        //ONLY_NUMBER=���ֶ�ֻ��Ϊ����,
        //FREETABLE_WIDTH_NOT_ALLOWED=����ָ�������Ԫ�����ʱ���������������ֵ,
        //FREETABLE_CELL_SPACING_ERROR=ֻ������������Ԫ����ʱ���˷�������Ч,
        //FREETABLE_LEFT_INDENT_ERROR=ֻ�е�������뷽ʽΪ����벢���ı���������Ϊ��ʱ���˷�������Ч,
        //FREETABLE_LOCATION_ERROR=������Ƕ���ı����ֻ���Ϊ��ʱ���������ñ���Ķ�λ��Ϣ,
        //FREETABLE_INCLUDE_VERTICAL_MERGE_CELL=�޷����ʴ˼����е������У���Ϊ������������ϲ��ĵ�Ԫ��,
        //FREETABLE_INCLUDE_MIXED_CELL=�޷����ʴ˼����е������У���Ϊ�������л�ϵĵ�Ԫ�����,
        //FREETABLE_CANNOT_MERGE_CELL=��ǰ��Ԫ�񼯺ϲ���һ���������򣬲��ܽ��кϲ���Ԫ�����,
        //WORKBOOK_CONTAINS_ONE_DOC_OR_SHEET_ATLEAST=����ɾ���ù��������ĵ��������ļ�������Ҫ����һ���ɼ��Ĺ�������򿪵��ĵ�,
        //HEAD_TAIL_EXIST_SPACE=��β���ڿո�,
        //DOCUMENTFIELD_NAME_INVALID=�����������к��зǷ��ַ�,
        //DOCUMENTFIELD_SAME_NAME=�����������Ѵ��ڣ�����������,
        //PRINTER_NOT_EXIST=�����ڴ�ӡ����,
        //PRINTER_NOT_CONNECTED=��ӡ���޷����ӣ�,
        //ERROR_RANGE_BY_NAME=δ�ҵ���ָ��������������Ŀ,
        //ERROR_RANGE_BY_INDEX=����������ָ������е�����,
        //ERROR_INLINE_SHAPE=������Ч,
        //ERROR_CAPTION_LABEL_NOEXIT=ָ������ע��ǩ������ ,
        //ERROR_AUTOCAPTION_NOEXIT=ָ�����Զ���ע������,
        //NO_OPENED_DOCUMENT=��Ϊû�д򿪵��ĵ���������һ������Ч��,
        //DOCUMENTFIELD_ONLY_READ=�������ڴ������ݺ��������ֻ��,
        //DOCUMENTFIELD_READONLY=������Ϊֻ��״̬���޷��༭

    }
    export const enum DocumentException {
        //INSERT_STRING_FAILURE=�����ַ�ʧ��!,
        //DELETE_STRING_FAILURE=ɾ���ַ�ʧ��!,
        //OFFSET_ERROR=��ʼβע����Ϊ����!,
        //DOCUMENT_NO_FOOTNOTE=���½�û�н�ע,
        //FOOTNOTE_ERROR1=����Խ�硣���½ڹ��� [0,,
        //FOOTNOTE_ERROR2=] ����ע��û���ҵ�����ָ���ĵ� ,
        //FOOTNOTE_ERROR3= ����ע��,
        //DOCUMENT_NO_ENDNOTE=���½�û��βע,
        //ENDNOTE_ERROR1=����Խ�硣���½ڹ��� [0,,
        //ENDNOTE_ERROR2=] ��βע��û���ҵ�����ָ���ĵ� ,
        //ENDNOTE_ERROR3= ��βע��,
        //OFFSET_IS_MINUS=�ַ�λ�ò���Ϊ����,
        //OFFSET_IS_OVER=�ַ�λ���Ѿ��������ĵ����ȡ�(�ĵ����� = ,
        //TABLEE_ERROR=û������ָ���ı���,
        //SSTABLEE_ERROR=û������ָ���Ĺ�����,
        //SHURU_LENGTH=������Ĺ��λ��,
        //DAYU_DOCLENGTH=���ڸ��½ڵ��ı�����,
        //INSERT_FIELD_TYPE_RANGE=���������� [0,,
        //INSERT_FIELD_NAME_RANGE=������и����������� [0,,
        //INSERT_FIELD_FORMAT_RANGE=����������������ʽ������ [0,,
        //BOOKNAME_ERROR1=��ǩ���ֲ���Ϊ�ջ�ո�,
        //BOOKNAME_ERROR2=��ǩ���ֿ�ͷ��һ����ĸ����Ϊ���� ''0-9'' �� ''_'',
        //BOOKNAME_ERROR3=��ǩ�����к��зǷ��ַ�,
        //TABSTOPS_ERROR1=����Խ�硣�ö����Ʊ�λ�� [0,,
        //TABSTOPS_ERROR2=] ��,������Ϸ����Ʊ�λ�����š�,
        //SECTION_ERROR1=����Խ�硣ȫ�Ĺ��� [0,,
        //SECTION_ERROR2=] ���½ڣ�û���ҵ�����ָ���ĵ� ,
        //SECTION_ERROR3= ���½�,
        //FILENAME_ERROR1=�ļ�������Ϊ�ա�,
        //FOOTNOTE_CANNOT_INSERT_OBJECT=��עβע���ܲ���ͼƬ�������ļ���Ӱ���ļ���,
        //FOOTNOTE_NO_AUTOFORMAT=��עβעû���Զ���ʽ���Ĺ���,
        //CANNOTPASTELINK=������ճ�����ӵ�����,
        //NOCOMPONENT=û����ָ�������,
        //NODATETIME=û����ָ����ʱ�����ڸ�ʽ,
        //NO_PHONETIC=����״̬����ȡƴ��,
        //HAS_STYLENAME=�Ѿ����ڣ����ѱ�����Ϊ������ʽ.,
        //NOHAS_STYLE=�����ڣ����߲����û��Զ�����ʽ������ɾ��,
        //NOHAS_STYLE1=��ʽ�����ڡ�ȡĬ����ʽ:,
        //HYPERLINK_NO_BOOKMARK=����ָ�����ĵ�û����ǩ,
        //HYPERLINK_NO_HEADING=����ָ�����ĵ�û�б���,
        //PROTECED_DOC=���ĵ����ڱ���״̬������ȡ���ĵ�����,
        //PROTECED_TRACKCHANGE=���ĵ����ڱ����޶�״̬������ȡ���ĵ�����,
        //UNPROTECED_DOC=���ĵ����ڲ�����״̬���ĵ��Ѿ��Ǳ��Ᵽ����,
        //NO_MERGE=ѡ�������Ǿ������򣬲��ܺϲ���Ԫ��������ѡ��,
        //NO_SPLIT=ѡ�������кϲ���Ԫ���Ҳ��Ǿ������򣬲��ܲ�ֵ�Ԫ��������ѡ��,
        //NOIN_TABLE=�����ڱ������ٴβ������������ѡ��������λ��,
        //NO_SHADING=���ܶ�ĳ���½ڻ�����ƪ�ĵ����õ���,
        //NOT_COLLAPSE_OR_EXPAND=�����ɱ����ڣ����ܶԶ�������۵���չ�������ı������,
        //NOT_MOVEUP_OR_DOWN=�����ɱ����ڣ����ܶԶ������ƻ�����,
        //NO_HEADANDFOOT=�����ĵ������ҳü��ҳ�Ź�����Ч,
        //SAME_NAME=�ñ༭�������Ѿ�������ͬ���ֵı���,
        //NO_TABLE_AS_NAME=û��ָ�����Ƶı���,
        //POSITION_OUTOF_TABLE=ָ����λ���ڱ����в�����,
        //UNFIT_ACTION_AREA=�����в���ִ�еĶ���ָ��,
        //VALUE_OUT_OF_BOUNDS=���������ѳ�����Χ,
        //DOC_EXIST=�ĵ��Ѿ�����,
        //NAME_INVALID=�ĵ������к��зǷ��ַ�,
        //NAME_TOO_LONG=�ĵ����ƹ�������಻���� 31 ���ַ�,
        //ERROR_FIELD_CODE=���ź�����벻����,
        //WP_NOT_OPEN=��ǰ���ִ���Ӧ�ó���δ������,
        //BARCODE_ERROR1=�����������Ϣ̫��,
        //BARCODE_ERROR2=����������������,
        //ERROR_MESSAGE3=�˷���ֻ��Ӧ�������Ĳ���,
        //ERROR_MESSAGE4=�˷���ֻ��Ӧ����ҳü��ҳ�Ż����Ĳ���,
        //FOOTENDNOT_ERROR1=��עβע���Զ����Ǳ���Ϊһ�ַ�ֵ,
        //FOOTENDNOTE_ERROR2=���ý�עβע�ı�Ÿ�ʽǰ����ʹ���Զ�����Ϊ��,
        //WBL_BULLET_ONLY=�˷���ֻ��Ӧ��������Ϊ��Ŀ���ŵ��б�����,
        //WBL_NUMBER_OUTLINENUMBER=�˷���ֻ��Ӧ��������Ϊ��Ż�༶��ŵ��б�����,
        //WBL_OUTLINENUMBER_ONLY=�˷���ֻ��Ӧ��������Ϊ�༶��ŵ��б�����,
        //SPREADSHEET_ONLY=�˷���ֻ��Ӧ���ڵ��ӱ�����,
        //WORDPROCESSOR_ONLY=�˷���ֻ��Ӧ�������ִ�����,
        //PRESENTATION_ONLY=�˷���ֻ��Ӧ���ڼ�������,
        //LETTER_STYLE_ONLY=ֻ�е���ֽ�ż�Ϊ��ֽ��ʽʱ�˷�������Ч,
        //GRID_LETTER_STYLE=�˷���ֻ��Ӧ���ڸ�ֽ�ż�Ϊ�������ֽ��ʽ,
        //COLUMN_WIDTH_NOTSET=��ǰ����Ϊ���������ʱ�������÷����е�������,
        //COLUMN_SPACE_NOTSET=��ǰ����Ϊ���������ʱ�������÷����е������,
        //GRID_SETTING_NONGRID=�˷�������Ӧ�����ĵ���������Ϊ��,
        //GRID_SETTING_LINESPACING1=�˷���ֻ��Ӧ���ڶ����о���������Сֵ���̶�ֵ��౶�о�,
        //PROTECTED_SECTION=�ýڴ��ڱ���״̬������ȡ������,
        //CAN_NOT_EDIT=ָ��λ�ô��ڲ��ɱ༭״̬,
        //FREETABLE_INSERT_HYPERLINK_ERROR=�����ɱ�����,�޷�ͬʱ�ڶ����Ԫ���в��볬����,
        //DOCUMENT_COUNT_ERROR=���ֻ�ܲ��� 64 ���ĵ�,
        //GRID_SETTING_NO_PAGEMARGE=ֻ���ڲ�ʹ��ҳ�߾�״̬�´˷�������Ч,
        //GRID_SETTING_HORIZONTAL=ֻ������ʾˮƽ�����״̬�´˷�������Ч,
        //GRID_SETTING_VERTICAL=ֻ������ʾ��ֱ�����״̬�´˷�������Ч,
        //CELL_IS_OVER=�ַ�λ���Ѿ������õ�Ԫ���ַ����ȡ�(��Ԫ���ַ����� = ,
        //WATERMARK_PICTURE=�˷���ֻ����ˮӡ����ΪͼƬˮӡʱ����Ч,
        //WATERMARK_TEXT=�˷���ֻ����ˮӡ����Ϊ����ˮӡʱ����Ч,
        //NO_BASESTYLE=��׼��ʽ������,
        //VALIDATE_DATE=��ǰ���ݱ���Ϊ��Ч����,
        //LESS_THAN_40=���ֶ������ֻ������ 40 ���ֽ�,
        //LESS_THAN_110=���ֶ������ֻ������ 110 ���ֽ�,
        //LESS_THAN_20=���ֶ������ֻ������ 20 ���ֽ�,
        //MORE_THAN_8=���ֶ����������� 8 ���ֽ�,
        //BARCODE_NUMBER=�������š�Ϊ��ѡ�ֶΣ�����������,
        //BARCODE_PRODUCE=������������λ��Ϊ��ѡ�ֶΣ�����������,
        //CROSSREF_ERROR=�������ô���,
        //LESS_THAN_30=���ֶ����ֻ������ 30 ���ֽ�,
        //LESS_THAN_150=���ֶ����ֻ������ 150 ���ֽ�,
        //FILE_FORMAT_INVALID=���ļ��ĸ�ʽ����ʶ��,
        //PAPER_DIMENSION_INVALID=ֽ�ųߴ�����������Ч,
        //FILE_DIRECTORY_NOT_EXIST=�ļ�Ŀ¼������,
        //LESS_THAN_8=���ֶβ��ܳ��� 8 ���ֽ�,
        //BOOKMARK_NOT_EXIST=�����ڴ���ǩ��,
        //ONLY_NUMBER=���ֶ�ֻ��Ϊ����,
        //FREETABLE_WIDTH_NOT_ALLOWED=����ָ�������Ԫ�����ʱ���������������ֵ,
        //FREETABLE_CELL_SPACING_ERROR=ֻ������������Ԫ����ʱ���˷�������Ч,
        //FREETABLE_LEFT_INDENT_ERROR=ֻ�е�������뷽ʽΪ����벢���ı���������Ϊ��ʱ���˷�������Ч,
        //FREETABLE_LOCATION_ERROR=������Ƕ���ı����ֻ���Ϊ��ʱ���������ñ���Ķ�λ��Ϣ,
        //FREETABLE_INCLUDE_VERTICAL_MERGE_CELL=�޷����ʴ˼����е������У���Ϊ������������ϲ��ĵ�Ԫ��,
        //FREETABLE_INCLUDE_MIXED_CELL=�޷����ʴ˼����е������У���Ϊ�������л�ϵĵ�Ԫ�����,
        //FREETABLE_CANNOT_MERGE_CELL=��ǰ��Ԫ�񼯺ϲ���һ���������򣬲��ܽ��кϲ���Ԫ�����,
        //WORKBOOK_CONTAINS_ONE_DOC_OR_SHEET_ATLEAST=����ɾ���ù��������ĵ��������ļ�������Ҫ����һ���ɼ��Ĺ�������򿪵��ĵ�,
        //HEAD_TAIL_EXIST_SPACE=��β���ڿո�,
        //DOCUMENTFIELD_NAME_INVALID=�����������к��зǷ��ַ�,
        //DOCUMENTFIELD_SAME_NAME=�����������Ѵ��ڣ�����������,
        //PRINTER_NOT_EXIST=�����ڴ�ӡ����,
        //PRINTER_NOT_CONNECTED=��ӡ���޷����ӣ�,
        //ERROR_RANGE_BY_NAME=δ�ҵ���ָ��������������Ŀ,
        //ERROR_RANGE_BY_INDEX=����������ָ������е�����,
        //ERROR_INLINE_SHAPE=������Ч,
        //ERROR_CAPTION_LABEL_NOEXIT=ָ������ע��ǩ������ ,
        //ERROR_AUTOCAPTION_NOEXIT=ָ�����Զ���ע������,
        //NO_OPENED_DOCUMENT=��Ϊû�д򿪵��ĵ���������һ������Ч��,
        //DOCUMENTFIELD_ONLY_READ=�������ڴ������ݺ��������ֻ��,
        //DOCUMENTFIELD_READONLY=������Ϊֻ��״̬���޷��༭

    }
}
