**English** | [Русский](MODULES_REFERENCE.md)
---

# MACROTools v2.0 — Complete Module Reference

> **Name:** MACROTools
> **Author:** VBATools
> **License:** Apache License
> **Type:** Excel VBA Add-in (.xlsb / .xlam)

---

## 📂 Add-in Core Modules

### `modAddinConst` — Constants
| Constant | Description |
|----------|-------------|
| `NAME_ADDIN` | Add-in name: "MACROTools" |
| `MENU_*` | VBE context menu constants |
| `TB_*` | Settings table names (SNIPETS, ABOUT, OPTIONS_IDEDENT, COMMENTS) |
| `FORMAT_DATE`, `FORMAT_TIME` | Date/time formats |
| `QUOTE_CHAR` | Quote character (`"`) |

---

### `modAddinCreateMenu` — VBE Context Menu Creation
| Procedure | Description |
|-----------|-------------|
| `Auto_Open()` | Automatically adds menus when add-in loads |
| `Auto_Close()` | Removes menus when add-in unloads |
| `RefreshMenu()` | Refreshes all VBE menus |
| `AddContextMenus()` | Creates all menu items: Move Controls, Tools, Code Window, MSForms |
| `AddButtom()` | Adds a button to VBE command bar |
| `AddComboBox()` | Adds ComboBox for scope selection (All VBAProject / Selected Module) |
| `DeleteContextMenus()` | Removes all created menus |
| `AddNewCommandBarMenu()` | Creates a new command bar menu |
| `DeleteButton()` | Removes a button by tag |

**Menu includes:**
- **MENU_MOVE_CONTROLS** — control movement (Up/Down/Left/Right)
- **MENU_TOOLS** — Hotkeys, FormatBuilder, MsgBoxBuilder, ProcedureBuilder, TODO, comments, formatting, Dim, Debug, snippets
- **MENU_CODE_WINDOW** — Swap [=], UPPER/lower case, Insert Code
- **MENU_FORMS** — alignment, renaming, control styles
- **MENU_PROJECT_WINDOW** — Copy Module
- **MENU_MS_FORMS** — styles, UPPER/lower case for forms

---

### `modAddinPubFun` — Public Functions (General)
| Procedure | Description |
|-----------|-------------|
| `Version()` | Returns add-in info (name, author, version, date) |
| `getCommentVBATools()` | ASCII art VBATools logo |
| `URLLinks()` | Opens URL in browser |
| `FileHave()` | Checks if file/folder exists |
| `sGetBaseName()` | Returns file name without extension |
| `sGetExtensionName()` | Returns file extension |
| `sGetFileName()` | Returns full file name |
| `sGetParentFolderName()` | Returns parent folder path |
| `MoveFile()` | Copies a file |
| `WorkbookIsOpen()` | Checks if workbook is open |
| `IsTableExists()` | Checks if ListObject exists on worksheet |
| `base64ToFile()` | Decodes Base64 string to file |
| `addTabelFromArray()` | Builds Markdown table from array |
| `fileDialogFun()` | File selection dialog |
| `GetFilesTable()` | Recursive folder traversal, returns file array |
| `GetArrayFromDictionary()` | Converts Dictionary to 2D array |
| `OutputResults()` | Outputs array to Excel worksheet with autofilter |
| `WriteErrorLog()` | Writes error to log (clsLogging) |
| `FilterArrayByText()` | Filters array by text in column |
| `GetTargetWorkbook()` | Shows target workbook selection form |

---

### `modAddinPubFunVBE` — Public VBE Functions
| Procedure | Description |
|-----------|-------------|
| `GetSelectControl()` | Returns selected control on UserForm designer |
| `getActiveCodePane()` | Returns active VBE code pane |
| `getActiveModule()` | Returns selected VBComponent |
| `GetCodeFromModule()` | Reads all code from module |
| `SetCodeInModule()` | Replaces all code in module |
| `SelectedLineColumnProcedure()` | Returns cursor position (line/column) |
| `WhatIsTextInComboBoxHave()` | Reads selected value from ComboBox menu |
| `TrimLinesTabAndSpase()` | Trim all module lines |
| `fnTrimLinesTabAndSpase()` | Trim lines (returns string) |
| `RemoveCommentsInVBACodeStrings()` | Removes comments from VBA code (handles strings) |
| `clearCodeStrings()` | Code cleanup: line breaks + comments + double empty lines |
| `FormatCodeToMultilineComma()` | Splits multi-line `Dim` by commas |
| `FormatSingleLineComma()` | Format single line with commas |
| `FormatCodeToMultilineColon()` | Splits lines by colon (`:`) |
| `FormatSingleLineColon()` | Format single line with colon |
| `RemoveBreaksLineInCode()` | Removes line continuations (` _`) |
| `GetProcedureDeclaration()` | Returns procedure declaration (with continuation handling) |
| `SingleSpace()` | Replaces multiple spaces with single space |
| `VBAIsTrusted()` | Checks VBA object model access |
| `GetProcedureName()` | Returns procedure name on line |
| `TryGetProcedureName()` | Attempts to get procedure name |
| `typeVariable()` | Determines type by suffix (`$`, `%`, `&`, `!`, `#`, `@`) |

---

### `modAddinPubFunVBEModule` — VBE Module Operations
| Procedure | Description |
|-----------|-------------|
| `exportModuleToFile()` | Exports module to file |
| `CopyModyleVBE()` | Copies module via context menu |
| `AddModuleToProject()` | Adds new module to project |
| `CopyModuleToProject()` | Copies module from one project to another |
| `CopyModuleTypeForm()` | Copies UserForm (via export/import) |
| `AddModuleUniqueName()` | Generates unique module name |
| `DeleteModuleToProject()` | Removes module from project |
| `getFileNameOnVBProject()` | Returns VBProject file name |
| `moduleLineCount()` | Counts code lines (excluding Option *) |
| `TypeProcedyre()` | Determines procedure type (Sub/Function/Property) |
| `TypeProcedyreModifier()` | Determines modifier (Public/Private) |
| `TypeExtensionModule()` | Returns extension by module type (.bas/.cls/.frm) |
| `moduleTypeName()` | Returns readable module type name |
| `getVBModuleByName()` | Returns module by name |

---

### `modAddinRibbonCallbacks` — Ribbon Callback Functions
Each procedure is bound to a Ribbon button:

| Callback | Action |
|----------|--------|
| `btnRefresh` | Refresh menus |
| `btnInfoFile` | Show file info |
| `btnOpenFile` | Unpack Office file |
| `btnCloseFile` | Pack Office file |
| `btnInToFile` | Archive file list |
| `btnUnProtectVBA` | Remove VBA project password |
| `btnExportVBA` | Module export manager |
| `btnUnProtectVBAUnivable` | Remove Unviewable protection |
| `btnProtectVBAUnivable` | Set Unviewable protection |
| `btnHiddenModule` | Hide modules |
| `btnUnProtectSheetsXML` | Remove sheet/workbook passwords |
| `btnObfuscator` | Obfuscate VBA project |
| `btnObfuscatorVariable` | Collect obfuscation data |
| `btnSerchVariableUnUsed` | Search unused variables |
| `btnAddStatisticAll` | Statistics: all |
| `btnAddStatisticForms` | Statistics: UserForms |
| `btnAddStatisticModules` | Statistics: modules |
| `btnAddStatisticDeclaretions` | Statistics: declarations |
| `btnAddStatisticProcedures` | Statistics: procedures |
| `btnAddStatisticShape` | Statistics: shapes |
| `btnParserLiterals` | Parse string literals |
| `btnReNameLiterals` | Rename literals |
| `btnToolCharMonitor` | Character monitor |
| `btnRegExpr` | Regular expression tester |
| `btnDeleteExternalLinks` | Delete external links |
| `btnReferenceStyle` | Toggle A1/R1C1 |
| `btnAddIn` | Add-in manager |
| `btnVBAWindowOpen` | Open VBE (Alt+F11) |
| `btnOptionsStyle` | Indent settings |
| `btnOptionsComment` | Comment settings |
| `btnBlackTheme` | Dark VBE theme |
| `btnWhiteTheme` | Light VBE theme |
| `btnOpenLogFile` | Open log file |
| `btnDeleteLogFile` | Clear logs |
| `btnAbout` | About |

---

### `modAddinThemeVBE` — VBE Themes
| Procedure | Description |
|-----------|-------------|
| `changeColorWhiteTheme()` | Enables light VBE theme |
| `changeColorDarkTheme()` | Enables dark VBE theme |
| `changeColorTheme()` | Writes colors to registry (`HKEY_CURRENT_USER\Software\Microsoft\VBA\`) |
| `GetVersionVBE()` | Returns VBE version for registry path |

---

### `modAddinInstall` — Add-in Installation
| Procedure | Description |
|-----------|-------------|
| `InstallationAddinMacroTools()` | Installs add-in to `Application.UserLibraryPath` folder |
| `ReadTableDataIntoTBArray()` | Reads settings tables from source file |
| `UpdateTablesFromTBArray()` | Updates settings tables in target workbook |

---

## 📂 File Operation Modules

### `modFilePassVBA` — VBA Password Removal
| Procedure | Description |
|-----------|-------------|
| `unProtectVBA()` | Removes password protection from VBA project |
| `unProtectVBAProjects()` | Hooks `DialogBoxParamA` function to bypass password dialog |
| `MyDialogBoxParam()` | Substituted dialog function (returns OK without password) |

> ⚠️ Uses API hooks (`VirtualProtect`, `MoveMemory`, `GetProcAddress`)

---

### `modFilePassVBAHideModule` — Module Hiding
| Procedure | Description |
|-----------|-------------|
| `hideModules()` | Hides modules from VBE project window (modifies `vbaProject.bin`) |
| `arrayByteJoin()` | Converts byte array to delimited string |
| `arrStringToByte()` | Converts string array to bytes |
| `addEmptyString()` | Generates placeholder of required length |

---

### `modFilePassVBAUnviewableDel` — Remove Unviewable Protection
| Procedure | Description |
|-----------|-------------|
| `delProtectVBAUnviewable()` | Removes "Unviewable VBA Project" protection |
| `unProtectVBAUnviewable()` | Replaces `CMG=`, `DPB=`, `GC=` keys with `CMC=`, `DPC=`, `CC=` |

---

### `modFilePassVBAUnviewableSet` — Set Unviewable Protection
| Procedure | Description |
|-----------|-------------|
| `setProtectVBAUnviewable()` | Sets "Unviewable VBA Project" protection |
| `ProtectVBAUnviewable()` | Injects "salt" from repeating CMG/DPB/GC keys |
| `addSaltString()` | Generates salt string of specified length |

---

### `modFilePassWBook` — Sheet/Workbook Password Removal
| Procedure | Description |
|-----------|-------------|
| `delPasswordWBook()` | Removes sheet and workbook structure protection via XML modification |

---

### `modFileProperty` — File Properties
| Procedure | Description |
|-----------|-------------|
| `GetOneProp()` | Returns single built-in document property |
| `getFilePropertiesCustomList()` | Custom properties array |
| `getFilePropertiesList()` | Built-in properties array |
| `addFilePropertyCustom()` | Adds custom property |
| `addFileProperty()` | Modifies built-in property |
| `delFilePropertiesCustomAll()` | Removes all custom properties |
| `delFilePropertiesAll()` | Clears all built-in properties |
| `delFilePropertyCustom()` | Removes single custom property |

---

### `modFileZipUnZip` — Archive/Unarchive
| Procedure | Description |
|-----------|-------------|
| `UnZipFile()` | Unpacks Office file to folder |
| `ZipFile()` | Packs folder back to Office file |
| `addListInFileFiles()` | Creates sheet with archive file list |
| `ZipAllFilesInFolder()` | Packs folder into ZIP |
| `FileUnZip()` | Unpacks ZIP to folder |
| `CreateEmptyZipFile()` | Creates empty ZIP file (PK signature) |
| `CopyItemsShell()` | Copy via Shell.Application |
| `DeleteFolderSafe()` | Safe folder deletion (via cmd) |

---

## 📂 String Literal Modules

### `modLiteralsGetCode` — Parse Literals from Code
| Procedure | Description |
|-----------|-------------|
| `parserLiteralsFormCode()` | Extracts string literals from VBA code (into Dictionary) |
| `ExtractQuotedStrings()` | Extracts quoted strings, handling `""` escaping |

---

### `modLiteralsGetMain` — Main Literal Parser
| Procedure | Description |
|-----------|-------------|
| `getAllLiteralsFile()` | Collects all literals: UserForms, VBA code, Ribbon UI |

**Constants:**
- `STR_UF` — UserForm literals sheet
- `STR_CODE` — Code literals sheet
- `STR_UI` — UI literals sheet

---

### `modLiteralsGetUI` — Parse Ribbon UI Literals
| Procedure | Description |
|-----------|-------------|
| `parserLiteralsFormUI()` | Extracts text from customUI/customUI14 XML |
| `parserLiteralsFormUIOnlyProcedures()` | Callback procedures only from Ribbon |
| `ProcessXMLPart()` | Process XML part (customUI) |
| `getLitersFromXML()` | Recursive XML tree traversal |
| `getLitersFromXMLNode()` | Extract node attributes |

---

### `modLiteralsGetUserForm` — Parse UserForm Literals
| Procedure | Description |
|-----------|-------------|
| `parserLiteralsFormControls()` | Extracts UserForm control texts (Caption, Value, ControlTipText) |
| `ProcessControl()` | Process single control (including MultiPage/TabStrip) |
| `GetPropertySafe()` | Safe property reading |
| `AddItemToDictionary()` | Add item to Dictionary |

---

### `modLiteralsSetCode` — Rename Literals in Code
| Procedure | Description |
|-----------|-------------|
| `renameLiteralsToCode()` | Replaces string literals in VBA code by array |
| `SaveModuleCode()` | Saves modified code to module |

---

### `modLiteralsSetMain` — Main Literal Renaming
| Procedure | Description |
|-----------|-------------|
| `ReNameLiteralsFile()` | Renames literals in all three areas (Code, UF, UI) |
| `loadArrayToSheet()` | Loads array to worksheet |
| `getArrayFromSheet()` | Reads array from worksheet |

---

### `modLiteralsSetUI` — Rename UI Literals
| Procedure | Description |
|-----------|-------------|
| `renameLiteralsToUI()` | Modifies Ribbon UI XML (attributes, ID) |
| `WriteXML()` | Writes changes to XML node |
| `ChangeAttribute()` | Changes XML node attribute |

**Enum `UIColumns`:** array columns for UI (ModuleType, XMLNodeName, TagName, IdOriginal, IdNew, AttrName, AttrText, AttrTextNew, Status)

---

### `modLiteralsSetUserForm` — Rename UserForm Literals
| Procedure | Description |
|-----------|-------------|
| `renameLiteralsToUserForm()` | Modifies UserForm control properties by array |
| `UpdateFormProperty()` | Updates form Caption itself |
| `UpdateObjectProperty()` | Updates control property |
| `UpdateNestedControl()` | Updates nested controls (Tab/Page) |
| `setValueInControl()` | Sets value via `CallByName` |

---

## 📂 Code Tools

### `modToolsLineIndent` — Automatic Indents (Smart Indenter)
| Procedure | Description |
|-----------|-------------|
| `RebuildModule()` | Formats indents in module/procedure/project |
| `RebuildCodeArray()` | Processes line array for formatting |
| `fnFindFirstItem()` | Finds first structured element |
| `CheckLine()` | Analyzes code line |
| `fnAlignFunction()` | Aligns continuation lines |

> Based on **Smart Indenter** by Stephen Bullen (Office Automation Ltd.)

---

### `modToolsObfuscation` — VBA Obfuscation
| Procedure | Description |
|-----------|-------------|
| `ObfuscationVBAProject()` | Runs obfuscation of selected VBA project via `clsObfuscator` |

**Constant:**
- `ms_VARIABLE_SHEET` — obfuscation variables sheet name

---

### `modToolsStatCode` — Code Statistics
| Procedure | Description |
|-----------|-------------|
| `addStatAll()` | Statistics: all |
| `addStatModules()` | Statistics: modules |
| `addStatModuleProcedures()` | Statistics: procedures |
| `addStatUserFormsControl()` | Statistics: UserForm controls |
| `addStatDeclaration()` | Statistics: declarations |
| `addListVariableProjectOfuscation()` | Collect obfuscation data |
| `RunStatCollection()` | Universal statistics collection |

**Enum `StatMode`:** msAll, msModules, msProcedures, msUserForms, msDeclarations

---

### `modToolsStatShape` — Shape Statistics
| Procedure | Description |
|-----------|-------------|
| `addShapeStatistic()` | Collects all shapes on sheets: name, text, macro |

Creates `SHAPES_VBA` sheet with hyperlinks to shapes.

---

### `modToolsAddComments` — Add Comments
| Procedure | Description |
|-----------|-------------|
| `sysAddHeaderTop()` | Inserts header comment in procedure |
| `sysAddModifiedTop()` | Inserts "Modified" line |
| `sysAddTODOTop()` | Inserts TODO comment |
| `addStringDelimetr()` | Separator line of `*` |
| `addArrFromTBComments()` | Comment template array from table |
| `TypeProcedyreComments()` | Procedure type format for comment |
| `TypeModuleComments()` | Module type format for comment |
| `GetCurrentProcInfo()` | Determines current procedure and position |
| `AddStringParamertFromProcedureDeclaration()` | Generates procedure parameter template |

---

### `modToolsDebugOnOff` — Debug.Print Toggle
| Procedure | Description |
|-----------|-------------|
| `debugOn()` | Uncomments `Debug.Print` |
| `debugOff()` | Comments `Debug.Print` |
| `findeReplaceWordInCodeVBPrj()` | Find/replace in entire project or module |
| `findeReplaceWordInCode()` | Find/replace in single module |

---

### `modToolsDelBreaksLine` — Remove Line Continuations
| Procedure | Description |
|-----------|-------------|
| `delBreaksLinesInCodeVBA()` | Removes ` _` (line continuation) from code |

---

### `modToolsDelCommentsInCode` — Remove Comments
| Procedure | Description |
|-----------|-------------|
| `delCommentsInCodeVBA()` | Removes all comments from VBA code |

---

### `modToolsDeleteLinksFile` — Remove External Links
| Procedure | Description |
|-----------|-------------|
| `ExternalLinkUtility()` | Scans file for external links |
| `ReportExternalLinks()` | Full report on all link types |
| `CheckCellFormulas()` | Links in cell formulas |
| `CheckShapeLinks()` | Links in shapes/objects |
| `CheckConditionalFormatting()` | Links in conditional formatting |
| `CheckChartLinks()` | Links in chart data sources |
| `CheckPivotTableLinks()` | Links in pivot table sources |
| `CheckDataValidationLinks()` | Links in data validation |
| `CheckNamedRangeLinks()` | Links in named ranges |
| `OutputLinkInfo()` | Outputs link info to report |

---

### `modToolsDelTwoEmptyStrings` — Remove Double Empty Lines
| Procedure | Description |
|-----------|-------------|
| `delTwoEmptyStrings()` | Removes series of empty lines (leaves one) |
| `delEmptyTwoString()` | Process single module |
| `deleteTwoEmptyCodeStrings()` | Returns cleaned code |

---

### `modToolsDimOneLine` — Dim Formatting
| Procedure | Description |
|-----------|-------------|
| `dimMultiLine()` | Splits `Dim` into multiple lines |
| `dimOneLine()` | Merges multiple `Dim` into one line |

---

### `modToolsLineNumbers` — Line Numbers
| Procedure | Description |
|-----------|-------------|
| `AddLineNumbersVBProject()` | Adds line numbers to procedures |
| `RemoveLineNumbersVBProject()` | Removes line numbers |
| `AddLineNumbersModule()` | Process single module |
| `RemoveLineNumbersModule()` | Remove numbers from module |
| `RemoveLineNumbers()` | Remove number from single line |
| `IsSelectCase()` | Check for `Select Case` |
| `IsMultiLineString()` | Check for line continuation (` _`) |
| `IsProcEndLine()` | Check for `End Sub/Function/Property` |
| `IsProcStartLine()` | Check for procedure declaration |

---

### `modToolsOptionsModule` — Module Options (Option *)
| Procedure | Description |
|-----------|-------------|
| `subOptionsForm()` | Option directive selection dialog |
| `insertOptionsExplicitAndPrivateModule()` | Quick insert `Option Explicit` + `Option Private Module` |
| `addString()` | Inserts Option at module start (replacing duplicates) |

**Supported directives:**
- `Option Explicit`
- `Option Private Module`
- `Option Compare Text`
- `Option Base 1`
- `Private Const MODULE_NAME As String = "..."`

---

### `modToolsRegExp` — Regular Expressions
| Procedure | Description |
|-----------|-------------|
| `RegExpStart()` | Start regex test |
| `RegExpGetMatches()` | Find and output all matches |
| `RegExpEnjoyReplace()` | Replace by regex |
| `RegExpFindReplace()` | Replace function (returns string) |
| `RegExpExecuteCollection()` | Returns match collection |
| `RegExpClearCells*()` | Clear worksheet cells |
| `AddSheetTestRegExp()` | Copies RegExp template to active sheet |

---

### `modToolsSnipets` — Code Snippets
| Procedure | Description |
|-----------|-------------|
| `InsertCodeFromSnippet()` | Inserts snippet by keyword |
| `AddSnippetEnumModule()` | Creates SNIPPETS module with Enum description |
| `DeleteSnippetEnumModule()` | Removes SNIPPETS module |
| `addSnipetModules()` | Adds modules/classes from snippet |
| `addSnipetForms()` | Adds UserForm from snippet |
| `getCodeFromShape()` | Reads code from shape on shSettings sheet |
| `findeValueInTabel()` | Search snippet in table |
| `AddSpaceCode()` | Adds indent to inserted code |
| `getArrayTBSnipets()` | Reads snippet table |
| `AddEnumCode()` | Generates Enum code from snippet table |

---

### `modToolsSwapEgual` — Swap Assignment
| Procedure | Description |
|-----------|-------------|
| `SwapEgual()` | Swaps left and right sides of `=` |
| `SwapEgualText()` | Text processing: `x = y` → `y = x` |

---

### `modToolsUnUsedVar` — Unused Variable Search
| Procedure | Description |
|-----------|-------------|
| `showFormUnUsedVariable()` | Shows analysis form |
| `AnalyzeCodeVBProjectUnUsed()` | Full VBA project analysis |
| `FindUnusedItems()` | Find unused items |
| `CheckUnusedModules()` | Check unused modules (Forms, Classes) |
| `CheckUnusedDeclarations()` | Check unused declarations |
| `CheckUnusedCodeElements()` | Check unused variables/procedures |
| `IsProcedureEventHandler()` | Determines if procedure is event handler |
| `GetLinkedShapeMacros()` | Collects macros bound to shapes |
| `FindInAllModulesCode()` | Search text in all modules |
| `CountRegexMatches()` | Count regex matches |
| `GetCollection()` | Creates Collection-dictionary from array |
| `GetControlsLookupCollection()` | UserForm controls dictionary |
| `GetClassEventsLookupCollection()` | WithEvents variable dictionary |

---

### `modToolsOther` — Other Utilities
| Procedure | Description |
|-----------|-------------|
| `CloseAllWindowsVBE()` | Closes all VBE windows except active |
| `AddLegendHotKeys()` | Outputs hotkeys reference |
| `showMsgBoxGenerator()` | Opens MsgBox constructor |
| `showBilderFormat()` | Opens formatting constructor |
| `showBilderProcedure()` | Opens procedure constructor |
| `ShowTODOList()` | Shows TODO list |

---

### `modTest` — Test Module
| Procedure | Description |
|-----------|-------------|
| `test()` | Test: generates JSON from `clsToolsVBACodeStatistics` |
| `TXTAddIntoTXTFile()` | Writes text to file |

---

## 📂 Classes

### `clsAnchors` — UserForm Anchors
| Method/Property | Description |
|-----------------|-------------|
| `AnchorEdge` (Enum) | anchorNone, anchorTop, anchorBottom, anchorLeft, anchorRight |
| `AddControl()` | Adds control with anchors |
| `ResizeControls()` | Applies anchors on form resize |

---

### `clsLogging` — CSV Logger
| Method | Description |
|--------|-------------|
| `LogInfo()` | Write INFO level |
| `LogWarning()` | Write WARNING level |
| `LogError()` | Write ERROR level |
| `ShowLog()` | Opens log file |
| `ResetLogs()` | Clears logs |

**Levels (Enum `LOG_LEVEL`):** INFO, WARNING, ERROR

---

### `clsObfuscator` — VBA Obfuscator
| Method | Description |
|--------|-------------|
| `Execute()` | Runs full project obfuscation |
| `GenerateName()` | Generates obfuscated name |
| `EncodeVariables()` | Renames variables |
| `EncodeProcedures()` | Renames procedures |
| `EncodeModules()` | Renames modules |

**Constants:**
- `mc_VARIABLE_SHEET` — variable report sheet
- `mc_REPORT_SHEET` — obfuscation report sheet
- `mc_NAME_PREFIX` — obfuscated name prefix

---

### `clsOfficeArchiveManager` — Office Archive Manager
| Method | Description |
|--------|-------------|
| `Initialize()` | Initialize with Office file |
| `UnZipFile()` | Unpack archive |
| `ZipFilesInFolder()` | Pack archive |
| `getBinaryArrayVBAProject()` | Read `vbaProject.bin` |
| `putBinaryArrayVBAProject()` | Write `vbaProject.bin` |
| `delPasswordWBook()` | Remove workbook password |
| `delPasswordSheet()` | Remove sheet password |
| `getArraySheetsName()` | Sheet names array |
| `getXMLDOC()` | Load XML document |
| `readXMLFromFile()` | Read XML file |
| `writeXMLToFile()` | Write XML file |
| `GetSettings()` | Returns path/name by `SettingsValue` enum |

**Enum `SettingsValue`:** FileFolder, FileFullName, FileName, FolderUnzipped, FolderZip, FolderXl, ExlFileWorkBook, FileCustomUI, FileCustomUI14 etc.

---

### `clsSort2DArray` — 2D Array Sorting
| Method | Description |
|--------|-------------|
| `Sort()` | Sorts 2D array by column |

---

### `clsToolsVBACodeStatistics` — VBA Code Statistics
| Method | Description |
|--------|-------------|
| `getJSONCodeBase()` | Returns JSON with full project analysis |
| `addListProcs()` | Collects procedures |
| `addListModules()` | Collects modules |
| `addListDeclarations()` | Collects declarations (variables, constants, Enum, Type) |
| `addListControlsUserForms()` | Collects UserForm controls |
| `getArrayCodeBase()` | Returns results array |
| `reBootArrayCodeBase()` | Clears internal array |

**Enum `stdColVBA`:** stdTypeElement, stdModuleT, stdModuleName, stdProcName, stdProcType, stdProcModifier, stdProcDeclaration, stdProcLines, stdCode, stdProcVariable, stdModuleCount

---

### `clsVBECommandHandler` — VBE Command Handler
| Property | Description |
|----------|-------------|
| `cmdButton` (WithEvents) | CommandBarButton with Click event |

Handles clicks on VBE context menu buttons and runs macro via `Application.Run`.

---

### `shInstallation` — Installation
| Method | Description |
|--------|-------------|
| `Initialize()` | Initialize installation process |

---

### `shRegExp` — Regex Sheet
Service class for RegExp test template sheet.

---

### `shSettings` — Settings
Service class for settings sheet. Contains tables:
- `TB_ABOUT` — add-in info
- `TB_SNIPETS` — code snippets
- `TB_OPTIONS_IDEDENT` — indent settings
- `TB_COMMENTS` — comment templates
- `TB_HOT_KEYS` — hotkeys

---

### `ThisWB` — Workbook Events
Event class for add-in workbook (`VB_PredeclaredId = True`).

---

## 📂 UserForms (18 forms)

| Form | Description |
|------|-------------|
| `frmAboutInfo` | Add-in info (version, author, license) |
| `frmBilderFormat` | Formatting constructor |
| `frmBilderMsgBoxGenerator` | MsgBox constructor |
| `frmBilderProcedure` | Procedure constructor |
| `frmCharsMonitor` | Character monitor/table |
| `frmDelPaswortSheetBook` | Remove sheet and workbook passwords |
| `frmHideModule` | Hide modules from VBA project |
| `frmInfoFile` | File info (properties) |
| `frmInfoFileLastAutor` | File author info |
| `frmListWBOpen` | Open workbooks list |
| `frmMendgerVBAModules` | VBA module manager (export/import) |
| `frmOptionsModule` | Option directive settings |
| `frmSettingsIndent` | Indent settings |
| `frmSettingsKomments` | Comment settings |
| `frmTODO` | TODO comment list |
| `frmVariableUnUsed` | Unused variable search |

---

## 🏗️ Project Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                        Ribbon / Menu                            │
├─────────────────────────────────────────────────────────────────┤
│ modAddinRibbonCallbacks  │  modAddinCreateMenu                  │
├─────────────────────────────────────────────────────────────────┤
│                     Core (PubFun)                               │
├─────────────────────────────────────────────────────────────────┤
│ modAddinPubFun  │  modAddinPubFunVBE  │  modAddinPubFunVBEModule│
├─────────────────────────────────────────────────────────────────┤
│                      Tools                                      │
├──────────────────┬──────────────────┬───────────────────────────┤
│ Code Operations  │ File Operations  │ Analysis & Statistics     │
│ LineIndent       │ FilePass*        │ StatCode                  │
│ Obfuscation      │ FileProperty     │ StatShape                 │
│ DebugOnOff       │ FileZipUnZip     │ UnUsedVar                 │
│ DelComments      │                  │ RegExp                    │
│ DelBreaksLine    │                  │ Snipets                   │
│ DelTwoEmpty      │                  │                           │
│ DimOneLine       │                  │                           │
│ LineNumbers      │                  │                           │
│ OptionsModule    │                  │                           │
│ SwapEgual        │                  │                           │
├──────────────────┴──────────────────┴───────────────────────────┤
│                    String Literals                              │
├─────────────────────────────────────────────────────────────────┤
│ Get: modLiteralsGet*  │  Set: modLiteralsSet*                   │
├─────────────────────────────────────────────────────────────────┤
│                      Classes                                    │
├──────────────────┬──────────────────┬───────────────────────────┤
│ clsAnchors       │ clsObfuscator    │ clsOfficeArchiveManager   │
│ clsLogging       │ clsSort2DArray   │ clsToolsVBACodeStatistics │
│ clsVBECommandHandler │ shSettings   │ ThisWB                    │
└─────────────────────────────────────────────────────────────────┘
```

---

*Reference generated: 03.04.2026*
