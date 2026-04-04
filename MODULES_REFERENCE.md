**Русский** | [English](MODULES_REFERENCE_ENG.md)
---

# MACROTools v2.0 — Полный справочник модулей

> **Название:** MACROTools  
> **Автор:** VBATools  
> **Лицензия:** Apache License  
> **Тип:** Excel VBA Add-in (.xlsb / .xlam)

---

## 📂 Модули ядра аддина

### `modAddinConst` — Константы
| Константа | Описание |
|-----------|----------|
| `NAME_ADDIN` | Имя аддина: "MACROTools" |
| `MENU_*` | Константы для контекстных меню VBE |
| `TB_*` | Имена таблиц настроек (SNIPETS, ABOUT, OPTIONS_IDEDENT, COMMENTS) |
| `FORMAT_DATE`, `FORMAT_TIME` | Форматы даты/времени |
| `QUOTE_CHAR` | Символ кавычки (`"`) |

---

### `modAddinCreateMenu` — Создание контекстных меню VBE
| Процедура | Описание |
|-----------|----------|
| `Auto_Open()` | Автоматически добавляет меню при загрузке аддина |
| `Auto_Close()` | Удаляет меню при выгрузке аддина |
| `RefreshMenu()` | Обновляет все меню VBE |
| `AddContextMenus()` | Создаёт все пункты меню: Move Controls, Tools, Code Window, MSForms |
| `AddButtom()` | Добавляет кнопку на панель команд VBE |
| `AddComboBox()` | Добавляет ComboBox для выбора области применения (All VBAProject / Selected Module) |
| `DeleteContextMenus()` | Удаляет все созданные меню |
| `AddNewCommandBarMenu()` | Создаёт новую командную панель |
| `DeleteButton()` | Удаляет кнопку по тегу |

**Меню включает:**
- **MENU_MOVE_CONTROLS** — перемещение контролов (Up/Down/Left/Right)
- **MENU_TOOLS** — Hotkeys, FormatBuilder, MsgBoxBuilder, ProcedureBuilder, TODO, комментарии, форматирование, Dim, Debug, сниппеты
- **MENU_CODE_WINDOW** — Swap [=], UPPER/lower case, Insert Code
- **MENU_FORMS** — выравнивание, переименование, стили контролов
- **MENU_PROJECT_WINDOW** — Copy Module
- **MENU_MS_FORMS** — стили, UPPER/lower case для форм

---

### `modAddinPubFun` — Публичные функции (общие)
| Процедура | Описание |
|-----------|----------|
| `Version()` | Возвращает информацию об аддине (имя, автор, версия, дата) |
| `getCommentVBATools()` | ASCII-арт логотип VBATools |
| `URLLinks()` | Открывает URL в браузере |
| `FileHave()` | Проверяет существование файла/папки |
| `sGetBaseName()` | Возвращает имя файла без расширения |
| `sGetExtensionName()` | Возвращает расширение файла |
| `sGetFileName()` | Возвращает полное имя файла |
| `sGetParentFolderName()` | Возвращает путь к родительской папке |
| `MoveFile()` | Копирует файл |
| `WorkbookIsOpen()` | Проверяет, открыта ли книга |
| `IsTableExists()` | Проверяет существование ListObject на листе |
| `base64ToFile()` | Декодирует Base64 строку в файл |
| `addTabelFromArray()` | Формирует Markdown-таблицу из массива |
| `fileDialogFun()` | Диалог выбора файлов |
| `GetFilesTable()` | Рекурсивный обход папки, возврат массива файлов |
| `GetArrayFromDictionary()` | Конвертирует Dictionary в 2D массив |
| `OutputResults()` | Выводит массив на лист Excel с автофильтром |
| `WriteErrorLog()` | Записывает ошибку в лог (clsLogging) |
| `FilterArrayByText()` | Фильтрует массив по тексту в столбце |
| `GetTargetWorkbook()` | Показывает форму выбора целевой книги |

---

### `modAddinPubFunVBE` — Публичные функции VBE
| Процедура | Описание |
|-----------|----------|
| `GetSelectControl()` | Возвращает выбранный контрол на UserForm дизайнере |
| `getActiveCodePane()` | Возвращает активную кодовую панель VBE |
| `getActiveModule()` | Возвращает выбранный VBComponent |
| `GetCodeFromModule()` | Читает весь код из модуля |
| `SetCodeInModule()` | Заменяет весь код в модуле |
| `SelectedLineColumnProcedure()` | Возвращает позицию курсора (строка/колонка) |
| `WhatIsTextInComboBoxHave()` | Читает выбранное значение из ComboBox меню |
| `TrimLinesTabAndSpase()` | Trim всех строк модуля |
| `fnTrimLinesTabAndSpase()` | Trim строк (возвращает строку) |
| `RemoveCommentsInVBACodeStrings()` | Удаляет комментарии из VBA-кода (учитывает строки) |
| `clearCodeStrings()` | Очистка кода: разрывы строк + комментарии + двойные пустые |
| `FormatCodeToMultilineComma()` | Разбивает многострочные `Dim` по запятым |
| `FormatSingleLineComma()` | Форматирование одной строки с запятыми |
| `FormatCodeToMultilineColon()` | Разбивает строки по двоеточию (`:`) |
| `FormatSingleLineColon()` | Форматирование одной строки с двоеточием |
| `RemoveBreaksLineInCode()` | Удаляет переносы строк (` _`) |
| `GetProcedureDeclaration()` | Возвращает декларацию процедуры (с обработкой continuation) |
| `SingleSpace()` | Заменяет множественные пробелы на один |
| `VBAIsTrusted()` | Проверяет доступ к объектной модели VBA |
| `GetProcedureName()` | Возвращает имя процедуры на строке |
| `TryGetProcedureName()` | Пытается получить имя процедуры |
| `typeVariable()` | Определяет тип по суффиксу (`$`, `%`, `&`, `!`, `#`, `@`) |

---

### `modAddinPubFunVBEModule` — Работа с модулями VBE
| Процедура | Описание |
|-----------|----------|
| `exportModuleToFile()` | Экспорт модуля в файл |
| `CopyModyleVBE()` | Копирование модуля через контекстное меню |
| `AddModuleToProject()` | Добавляет новый модуль в проект |
| `CopyModuleToProject()` | Копирует модуль из одного проекта в другой |
| `CopyModuleTypeForm()` | Копирование UserForm (через экспорт/импорт) |
| `AddModuleUniqueName()` | Генерирует уникальное имя модуля |
| `DeleteModuleToProject()` | Удаляет модуль из проекта |
| `getFileNameOnVBProject()` | Возвращает имя файла VBProject |
| `moduleLineCount()` | Считает строки кода (исключая Option *) |
| `TypeProcedyre()` | Определяет тип процедуры (Sub/Function/Property) |
| `TypeProcedyreModifier()` | Определяет модификатор (Public/Private) |
| `TypeExtensionModule()` | Возвращает расширение по типу модуля (.bas/.cls/.frm) |
| `moduleTypeName()` | Возвращает читаемое имя типа модуля |
| `getVBModuleByName()` | Возвращает модуль по имени |

---

### `modAddinRibbonCallbacks` — Callback-функции Ribbon
Каждая процедура привязана к кнопке на Ribbon:

| Callback | Действие |
|----------|----------|
| `btnRefresh` | Обновить меню |
| `btnInfoFile` | Показать информацию о файле |
| `btnOpenFile` | Распаковать Office файл |
| `btnCloseFile` | Запаковать Office файл |
| `btnInToFile` | Список файлов архива |
| `btnUnProtectVBA` | Снять пароль с VBA проекта |
| `btnExportVBA` | Менеджер экспорта модулей |
| `btnUnProtectVBAUnivable` | Удалить Unviewable защиту |
| `btnProtectVBAUnivable` | Установить Unviewable защиту |
| `btnHiddenModule` | Скрыть модули |
| `btnUnProtectSheetsXML` | Удалить пароли листов/книги |
| `btnObfuscator` | Обфускация VBA проекта |
| `btnObfuscatorVariable` | Сбор данных для обфускации |
| `btnSerchVariableUnUsed` | Поиск неиспользуемых переменных |
| `btnAddStatisticAll` | Статистика: всё |
| `btnAddStatisticForms` | Статистика: UserForms |
| `btnAddStatisticModules` | Статистика: модули |
| `btnAddStatisticDeclaretions` | Статистика: декларации |
| `btnAddStatisticProcedures` | Статистика: процедуры |
| `btnAddStatisticShape` | Статистика: шейпы |
| `btnParserLiterals` | Парсинг строковых литералов |
| `btnReNameLiterals` | Переименование литералов |
| `btnToolCharMonitor` | Монитор символов |
| `btnRegExpr` | Тестер регулярных выражений |
| `btnDeleteExternalLinks` | Удаление внешних ссылок |
| `btnReferenceStyle` | Переключение A1/R1C1 |
| `btnAddIn` | Менеджер надстроек |
| `btnVBAWindowOpen` | Открыть VBE (Alt+F11) |
| `btnOptionsStyle` | Настройки отступов |
| `btnOptionsComment` | Настройки комментариев |
| `btnBlackTheme` | Тёмная тема VBE |
| `btnWhiteTheme` | Светлая тема VBE |
| `btnOpenLogFile` | Открыть лог-файл |
| `btnDeleteLogFile` | Очистить логи |
| `btnAbout` | О программе |

---

### `modAddinThemeVBE` — Темы VBE
| Процедура | Описание |
|-----------|----------|
| `changeColorWhiteTheme()` | Включает светлую тему VBE |
| `changeColorDarkTheme()` | Включает тёмную тему VBE |
| `changeColorTheme()` | Записывает цвета в реестр (`HKEY_CURRENT_USER\Software\Microsoft\VBA\`) |
| `GetVersionVBE()` | Возвращает версию VBE для пути реестра |

---

### `modAddinInstall` — Установка аддина
| Процедура | Описание |
|-----------|----------|
| `InstallationAddinMacroTools()` | Устанавливает аддин в папку `Application.UserLibraryPath` |
| `ReadTableDataIntoTBArray()` | Читает таблицы настроек из исходного файла |
| `UpdateTablesFromTBArray()` | Обновляет таблицы настроек в целевой книге |

---

## 📂 Модули работы с файлами

### `modFilePassVBA` — Снятие пароля VBA
| Процедура | Описание |
|-----------|----------|
| `unProtectVBA()` | Снимает защиту паролем с VBA проекта |
| `unProtectVBAProjects()` | Hook функции `DialogBoxParamA` для обхода диалога пароля |
| `MyDialogBoxParam()` | Подменённая функция диалога (возвращает OK без ввода пароля) |

> ⚠️ Использует API-хуки (`VirtualProtect`, `MoveMemory`, `GetProcAddress`)

---

### `modFilePassVBAHideModule` — Скрытие модулей
| Процедура | Описание |
|-----------|----------|
| `hideModules()` | Скрывает модули из проектного окна VBE (модификация `vbaProject.bin`) |
| `arrayByteJoin()` | Конвертация байтового массива в строку с разделителем |
| `arrStringToByte()` | Конвертация массива строк в байты |
| `addEmptyString()` | Генерация заполнителя нужной длины |

---

### `modFilePassVBAUnviewableDel` — Удаление Unviewable защиты
| Процедура | Описание |
|-----------|----------|
| `delProtectVBAUnviewable()` | Удаляет защиту "Unviewable VBA Project" |
| `unProtectVBAUnviewable()` | Заменяет ключи `CMG=`, `DPB=`, `GC=` на `CMC=`, `DPC=`, `CC=` |

---

### `modFilePassVBAUnviewableSet` — Установка Unviewable защиты
| Процедура | Описание |
|-----------|----------|
| `setProtectVBAUnviewable()` | Устанавливает защиту "Unviewable VBA Project" |
| `ProtectVBAUnviewable()` | Внедряет "salt" из повторяющихся ключей CMG/DPB/GC |
| `addSaltString()` | Генерирует строку-"соль" заданной длины |

---

### `modFilePassWBook` — Удаление паролей листов/книги
| Процедура | Описание |
|-----------|----------|
| `delPasswordWBook()` | Снимает защиту листов и структуры книги через модификацию XML |

---

### `modFileProperty` — Свойства файлов
| Процедура | Описание |
|-----------|----------|
| `GetOneProp()` | Возвращает одно встроенное свойство документа |
| `getFilePropertiesCustomList()` | Массив пользовательских свойств |
| `getFilePropertiesList()` | Массив встроенных свойств |
| `addFilePropertyCustom()` | Добавляет пользовательское свойство |
| `addFileProperty()` | Изменяет встроенное свойство |
| `delFilePropertiesCustomAll()` | Удаляет все пользовательские свойства |
| `delFilePropertiesAll()` | Очищает все встроенные свойства |
| `delFilePropertyCustom()` | Удаляет одно пользовательское свойство |

---

### `modFileZipUnZip` — Архивация/разархивация
| Процедура | Описание |
|-----------|----------|
| `UnZipFile()` | Распаковывает Office файл в папку |
| `ZipFile()` | Запаковывает папку обратно в Office файл |
| `addListInFileFiles()` | Создаёт лист со списком всех файлов архива |
| `ZipAllFilesInFolder()` | Упаковывает папку в ZIP |
| `FileUnZip()` | Распаковывает ZIP в папку |
| `CreateEmptyZipFile()` | Создаёт пустой ZIP-файл (PK-сигнатура) |
| `CopyItemsShell()` | Копирование через Shell.Application |
| `DeleteFolderSafe()` | Надёжное удаление папки (через cmd) |

---

## 📂 Модули работы со строковыми литералами

### `modLiteralsGetCode` — Парсинг литералов из кода
| Процедура | Описание |
|-----------|----------|
| `parserLiteralsFormCode()` | Извлекает строковые литералы из кода VBA (в Dictionary) |
| `ExtractQuotedStrings()` | Извлекает строки в кавычках, обрабатывая экранирование `""` |

---

### `modLiteralsGetMain` — Главный парсер литералов
| Процедура | Описание |
|-----------|----------|
| `getAllLiteralsFile()` | Собирает все литералы: UserForms, код VBA, Ribbon UI |

**Константы:**
- `STR_UF` — лист для литералов UserForm
- `STR_CODE` — лист для литералов кода
- `STR_UI` — лист для литералов UI

---

### `modLiteralsGetUI` — Парсинг литералов Ribbon UI
| Процедура | Описание |
|-----------|----------|
| `parserLiteralsFormUI()` | Извлекает текст из customUI/customUI14 XML |
| `parserLiteralsFormUIOnlyProcedures()` | Только callback-процедуры Ribbon |
| `ProcessXMLPart()` | Обработка XML части (customUI) |
| `getLitersFromXML()` | Рекурсивный обход XML-дерева |
| `getLitersFromXMLNode()` | Извлечение атрибутов узла |

---

### `modLiteralsGetUserForm` — Парсинг литералов UserForm
| Процедура | Описание |
|-----------|----------|
| `parserLiteralsFormControls()` | Извлекает тексты контролов UserForm (Caption, Value, ControlTipText) |
| `ProcessControl()` | Обработка одного контрола (включая MultiPage/TabStrip) |
| `GetPropertySafe()` | Безопасное чтение свойства |
| `AddItemToDictionary()` | Добавление элемента в Dictionary |

---

### `modLiteralsSetCode` — Переименование литералов в коде
| Процедура | Описание |
|-----------|----------|
| `renameLiteralsToCode()` | Заменяет строковые литералы в коде VBA по массиву |
| `SaveModuleCode()` | Сохраняет изменённый код в модуль |

---

### `modLiteralsSetMain` — Главное переименование литералов
| Процедура | Описание |
|-----------|----------|
| `ReNameLiteralsFile()` | Переименовывает литералы во всех трёх областях (Code, UF, UI) |
| `loadArrayToSheet()` | Загружает массив на лист |
| `getArrayFromSheet()` | Читает массив из листа |

---

### `modLiteralsSetUI` — Переименование литералов UI
| Процедура | Описание |
|-----------|----------|
| `renameLiteralsToUI()` | Модифицирует XML Ribbon UI (атрибуты, ID) |
| `WriteXML()` | Записывает изменения в XML узел |
| `ChangeAttribute()` | Изменяет атрибут XML узла |

**Enum `UIColumns`:** столбцы массива UI (ModuleType, XMLNodeName, TagName, IdOriginal, IdNew, AttrName, AttrText, AttrTextNew, Status)

---

### `modLiteralsSetUserForm` — Переименование литералов UserForm
| Процедура | Описание |
|-----------|----------|
| `renameLiteralsToUserForm()` | Изменяет свойства контролов UserForm по массиву |
| `UpdateFormProperty()` | Обновляет Caption самой формы |
| `UpdateObjectProperty()` | Обновляет свойство контрола |
| `UpdateNestedControl()` | Обновляет вложенные контролы (Tab/Page) |
| `setValueInControl()` | Устанавливает значение через `CallByName` |

---

## 📂 Инструменты работы с кодом

### `modToolsLineIndent` — Автоматические отступы (Smart Indenter)
| Процедура | Описание |
|-----------|----------|
| `RebuildModule()` | Форматирует отступы в модуле/процедуре/проекте |
| `RebuildCodeArray()` | Обрабатывает массив строк для форматирования |
| `fnFindFirstItem()` | Находит первый структурированный элемент |
| `CheckLine()` | Анализирует строку кода |
| `fnAlignFunction()` | Выравнивает строки продолжения |

> На основе **Smart Indenter** by Stephen Bullen (Office Automation Ltd.)

---

### `modToolsObfuscation` — Обфускация VBA
| Процедура | Описание |
|-----------|----------|
| `ObfuscationVBAProject()` | Запускает обфускацию выбранного VBA проекта через `clsObfuscator` |

**Константа:**
- `ms_VARIABLE_SHEET` — имя листа с переменными обфускации

---

### `modToolsStatCode` — Статистика кода
| Процедура | Описание |
|-----------|----------|
| `addStatAll()` | Статистика: всё |
| `addStatModules()` | Статистика: модули |
| `addStatModuleProcedures()` | Статистика: процедуры |
| `addStatUserFormsControl()` | Статистика: контролы UserForm |
| `addStatDeclaration()` | Статистика: декларации |
| `addListVariableProjectOfuscation()` | Сбор данных для обфускации |
| `RunStatCollection()` | Универсальный запуск сбора статистики |

**Enum `StatMode`:** msAll, msModules, msProcedures, msUserForms, msDeclarations

---

### `modToolsStatShape` — Статистика шейпов
| Процедура | Описание |
|-----------|----------|
| `addShapeStatistic()` | Собирает все шейпы на листах: имя, текст, макрос |

Создаёт лист `SHAPES_VBA` с гиперссылками на шейпы.

---

### `modToolsAddComments` — Добавление комментариев
| Процедура | Описание |
|-----------|----------|
| `sysAddHeaderTop()` | Вставляет шапку-комментарий в процедуру |
| `sysAddModifiedTop()` | Вставляет строку "Modified" |
| `sysAddTODOTop()` | Вставляет TODO-комментарий |
| `addStringDelimetr()` | Строка-разделитель из `*` |
| `addArrFromTBComments()` | Массив шаблонов комментариев из таблицы |
| `TypeProcedyreComments()` | Формат типа процедуры для комментария |
| `TypeModuleComments()` | Формат типа модуля для комментария |
| `GetCurrentProcInfo()` | Определяет текущую процедуру и позицию |
| `AddStringParamertFromProcedureDeclaration()` | Генерирует шаблон параметров процедуры |

---

### `modToolsDebugOnOff` — Включение/отключение Debug.Print
| Процедура | Описание |
|-----------|----------|
| `debugOn()` | Раскомментирует `Debug.Print` |
| `debugOff()` | Комментирует `Debug.Print` |
| `findeReplaceWordInCodeVBPrj()` | Поиск/замена во всём проекте или модуле |
| `findeReplaceWordInCode()` | Поиск/замена в одном модуле |

---

### `modToolsDelBreaksLine` — Удаление переносов строк
| Процедура | Описание |
|-----------|----------|
| `delBreaksLinesInCodeVBA()` | Удаляет ` _` (продолжение строки) из кода |

---

### `modToolsDelCommentsInCode` — Удаление комментариев
| Процедура | Описание |
|-----------|----------|
| `delCommentsInCodeVBA()` | Удаляет все комментарии из кода VBA |

---

### `modToolsDeleteLinksFile` — Удаление внешних ссылок
| Процедура | Описание |
|-----------|----------|
| `ExternalLinkUtility()` | Сканирует файл на внешние ссылки |
| `ReportExternalLinks()` | Полный отчёт по всем типам ссылок |
| `CheckCellFormulas()` | Ссылки в формулах ячеек |
| `CheckShapeLinks()` | Ссылки в шейпах/объектах |
| `CheckConditionalFormatting()` | Ссылки в условном форматировании |
| `CheckChartLinks()` | Ссылки в источниках данных диаграмм |
| `CheckPivotTableLinks()` | Ссылки в источниках сводных таблиц |
| `CheckDataValidationLinks()` | Ссылки в проверке данных |
| `CheckNamedRangeLinks()` | Ссылки в именованных диапазонах |
| `OutputLinkInfo()` | Выводит информацию о ссылке в отчёт |

---

### `modToolsDelTwoEmptyStrings` — Удаление двойных пустых строк
| Процедура | Описание |
|-----------|----------|
| `delTwoEmptyStrings()` | Удаляет серии пустых строк (оставляет одну) |
| `delEmptyTwoString()` | Обработка одного модуля |
| `deleteTwoEmptyCodeStrings()` | Возвращает очищенный код |

---

### `modToolsDimOneLine` — Форматирование Dim
| Процедура | Описание |
|-----------|----------|
| `dimMultiLine()` | Разбивает `Dim` на несколько строк |
| `dimOneLine()` | Объединяет несколько `Dim` в одну строку |

---

### `modToolsLineNumbers` — Номера строк
| Процедура | Описание |
|-----------|----------|
| `AddLineNumbersVBProject()` | Добавляет номера строк в процедуры |
| `RemoveLineNumbersVBProject()` | Удаляет номера строк |
| `AddLineNumbersModule()` | Обработка одного модуля |
| `RemoveLineNumbersModule()` | Удаление номеров из модуля |
| `RemoveLineNumbers()` | Удаляет номер с одной строки |
| `IsSelectCase()` | Проверка на `Select Case` |
| `IsMultiLineString()` | Проверка на продолжение строки (` _`) |
| `IsProcEndLine()` | Проверка на `End Sub/Function/Property` |
| `IsProcStartLine()` | Проверка на декларацию процедуры |

---

### `modToolsOptionsModule` — Настройки модуля (Option *)
| Процедура | Описание |
|-----------|----------|
| `subOptionsForm()` | Диалог выбора Option директив |
| `insertOptionsExplicitAndPrivateModule()` | Быстрая вставка `Option Explicit` + `Option Private Module` |
| `addString()` | Вставляет Option в начало модуля (заменяя дубликаты) |

**Поддерживаемые директивы:**
- `Option Explicit`
- `Option Private Module`
- `Option Compare Text`
- `Option Base 1`
- `Private Const MODULE_NAME As String = "..."`

---

### `modToolsRegExp` — Регулярные выражения
| Процедура | Описание |
|-----------|----------|
| `RegExpStart()` | Запуск теста регулярных выражений |
| `RegExpGetMatches()` | Поиск и вывод всех совпадений |
| `RegExpEnjoyReplace()` | Замена по регулярному выражению |
| `RegExpFindReplace()` | Функция замены (возвращает строку) |
| `RegExpExecuteCollection()` | Возвращает коллекцию совпадений |
| `RegExpClearCells*()` | Очистка ячеек листа |
| `AddSheetTestRegExp()` | Копирует шаблон RegExp на активный лист |

---

### `modToolsSnipets` — Сниппеты кода
| Процедура | Описание |
|-----------|----------|
| `InsertCodeFromSnippet()` | Вставляет сниппет по ключевому слову |
| `AddSnippetEnumModule()` | Создаёт модуль SNIPPETS с Enum-описанием |
| `DeleteSnippetEnumModule()` | Удаляет модуль SNIPPETS |
| `addSnipetModules()` | Добавляет модули/классы из сниппета |
| `addSnipetForms()` | Добавляет UserForm из сниппета |
| `getCodeFromShape()` | Читает код из фигуры на листе shSettings |
| `findeValueInTabel()` | Поиск сниппета в таблице |
| `AddSpaceCode()` | Добавляет отступ к вставляемому коду |
| `getArrayTBSnipets()` | Читает таблицу сниппетов |
| `AddEnumCode()` | Генерирует код Enum из таблицы сниппетов |

---

### `modToolsSwapEgual` — Swap присваивания
| Процедура | Описание |
|-----------|----------|
| `SwapEgual()` | Меняет местами левую и правую часть `=` |
| `SwapEgualText()` | Обработка текста: `x = y` → `y = x` |

---

### `modToolsUnUsedVar` — Поиск неиспользуемых переменных
| Процедура | Описание |
|-----------|----------|
| `showFormUnUsedVariable()` | Показывает форму анализа |
| `AnalyzeCodeVBProjectUnUsed()` | Полный анализ VBA проекта |
| `FindUnusedItems()` | Поиск неиспользуемых элементов |
| `CheckUnusedModules()` | Проверка неиспользуемых модулей (Forms, Classes) |
| `CheckUnusedDeclarations()` | Проверка неиспользуемых деклараций |
| `CheckUnusedCodeElements()` | Проверка неиспользуемых переменных/процедур |
| `IsProcedureEventHandler()` | Определяет, является ли процедура обработчиком событий |
| `GetLinkedShapeMacros()` | Собирает макросы, привязанные к шейпам |
| `FindInAllModulesCode()` | Ищет текст во всех модулях |
| `CountRegexMatches()` | Считает совпадения RegExp |
| `GetCollection()` | Создаёт Collection-словарь из массива |
| `GetControlsLookupCollection()` | Словарь контролов UserForm |
| `GetClassEventsLookupCollection()` | Словарь переменных WithEvents |

---

### `modToolsOther` — Прочие утилиты
| Процедура | Описание |
|-----------|----------|
| `CloseAllWindowsVBE()` | Закрывает все окна VBE, кроме активного |
| `AddLegendHotKeys()` | Выводит справку по горячим клавишам |
| `showMsgBoxGenerator()` | Открывает конструктор MsgBox |
| `showBilderFormat()` | Открывает конструктор форматирования |
| `showBilderProcedure()` | Открывает конструктор процедур |
| `ShowTODOList()` | Показывает список TODO |

---

### `modTest` — Тестовый модуль
| Процедура | Описание |
|-----------|----------|
| `test()` | Тест: генерирует JSON из `clsToolsVBACodeStatistics` |
| `TXTAddIntoTXTFile()` | Записывает текст в файл |

---

## 📂 Классы

### `clsAnchors` — Якоря для UserForm
| Метод/Свойство | Описание |
|----------------|----------|
| `AnchorEdge` (Enum) | anchorNone, anchorTop, anchorBottom, anchorLeft, anchorRight |
| `AddControl()` | Добавляет контрол с якорями |
| `ResizeControls()` | Применяет якоря при изменении размера формы |

---

### `clsLogging` — Логгер CSV
| Метод | Описание |
|-------|----------|
| `LogInfo()` | Запись INFO уровня |
| `LogWarning()` | Запись WARNING уровня |
| `LogError()` | Запись ERROR уровня |
| `ShowLog()` | Открывает лог-файл |
| `ResetLogs()` | Очищает логи |

**Уровни (Enum `LOG_LEVEL`):** INFO, WARNING, ERROR

---

### `clsObfuscator` — Обфускатор VBA
| Метод | Описание |
|-------|----------|
| `Execute()` | Запускает полную обфускацию проекта |
| `GenerateName()` | Генерирует обфусцированное имя |
| `EncodeVariables()` | Переименовывает переменные |
| `EncodeProcedures()` | Переименовывает процедуры |
| `EncodeModules()` | Переименовывает модули |

**Константы:**
- `mc_VARIABLE_SHEET` — лист отчёта переменных
- `mc_REPORT_SHEET` — лист отчёта обфускации
- `mc_NAME_PREFIX` — префикс обфусцированных имён

---

### `clsOfficeArchiveManager` — Менеджер архивов Office
| Метод | Описание |
|-------|----------|
| `Initialize()` | Инициализация с файлом Office |
| `UnZipFile()` | Распаковка архива |
| `ZipFilesInFolder()` | Упаковка архива |
| `getBinaryArrayVBAProject()` | Чтение `vbaProject.bin` |
| `putBinaryArrayVBAProject()` | Запись `vbaProject.bin` |
| `delPasswordWBook()` | Удаление пароля книги |
| `delPasswordSheet()` | Удаление пароля листа |
| `getArraySheetsName()` | Массив имён листов |
| `getXMLDOC()` | Загрузка XML документа |
| `readXMLFromFile()` | Чтение XML файла |
| `writeXMLToFile()` | Запись XML файла |
| `GetSettings()` | Возвращает путь/имя по enum `SettingsValue` |

**Enum `SettingsValue`:** FileFolder, FileFullName, FileName, FolderUnzipped, FolderZip, FolderXl, ExlFileWorkBook, FileCustomUI, FileCustomUI14 и др.

---

### `clsSort2DArray` — Сортировка 2D массивов
| Метод | Описание |
|-------|----------|
| `Sort()` | Сортирует двумерный массив по столбцу |

---

### `clsToolsVBACodeStatistics` — Статистика VBA кода
| Метод | Описание |
|-------|----------|
| `getJSONCodeBase()` | Возвращает JSON с полным анализом проекта |
| `addListProcs()` | Собирает процедуры |
| `addListModules()` | Собирает модули |
| `addListDeclarations()` | Собирает декларации (переменные, константы, Enum, Type) |
| `addListControlsUserForms()` | Собирает контролы UserForm |
| `getArrayCodeBase()` | Возвращает массив результатов |
| `reBootArrayCodeBase()` | Очищает внутренний массив |

**Enum `stdColVBA`:** stdTypeElement, stdModuleT, stdModuleName, stdProcName, stdProcType, stdProcModifier, stdProcDeclaration, stdProcLines, stdCode, stdProcVariable, stdModuleCount

---

### `clsVBECommandHandler` — Обработчик команд VBE
| Свойство | Описание |
|----------|----------|
| `cmdButton` (WithEvents) | CommandBarButton с событием Click |

Обрабатывает клики по кнопкам контекстного меню VBE и запускает макрос через `Application.Run`.

---

### `shInstallation` — Установка
| Метод | Описание |
|-------|----------|
| `Initialize()` | Инициализация процесса установки |

---

### `shRegExp` — Лист регулярных выражений
Служебный класс листа-шаблона для тестирования RegExp.

---

### `shSettings` — Настройки
Служебный класс листа настроек. Содержит таблицы:
- `TB_ABOUT` — информация об аддине
- `TB_SNIPETS` — сниппеты кода
- `TB_OPTIONS_IDEDENT` — настройки отступов
- `TB_COMMENTS` — шаблоны комментариев
- `TB_HOT_KEYS` — горячие клавиши

---

### `ThisWB` — События книги
Событийный класс книги аддина (`VB_PredeclaredId = True`).

---

## 📂 UserForms (32 формы)

| Форма | Описание |
|-------|----------|
| `frmAboutInfo` | Информация об аддине (версия, автор, лицензия) |
| `frmBilderFormat` | Конструктор форматирования строки |
| `frmBilderMsgBoxGenerator` | Конструктор MsgBox |
| `frmBilderProcedure` | Конструктор процедур |
| `frmCharsMonitor` | Монитор/таблица символов |
| `frmDelPaswortSheetBook` | Удаление паролей листов и книги |
| `frmHideModule` | Скрытие модулей из VBA проекта |
| `frmInfoFile` | Информация о файле (свойства) |
| `frmInfoFileLastAutor` | Информация об авторе файла |
| `frmListWBOpen` | Список открытых книг |
| `frmMendgerVBAModules` | Менеджер модулей VBA (экспорт/импорт) |
| `frmOptionsModule` | Настройки Option директив |
| `frmSettingsIndent` | Настройки отступов |
| `frmSettingsKomments` | Настройки комментариев |
| `frmTODO` | Список TODO комментариев |
| `frmVariableUnUsed` | Поиск неиспользуемых переменных |

---

## 🏗️ Архитектура проекта

```
┌─────────────────────────────────────────────────────────────────┐
│                        Ribbon / Меню                            │
├─────────────────────────────────────────────────────────────────┤
│ modAddinRibbonCallbacks  │  modAddinCreateMenu                  │
├─────────────────────────────────────────────────────────────────┤
│                     Ядро (PubFun)                               │
├─────────────────────────────────────────────────────────────────┤
│ modAddinPubFun  │  modAddinPubFunVBE  │  modAddinPubFunVBEModule│
├─────────────────────────────────────────────────────────────────┤
│                      Инструменты                                │
├──────────────────┬──────────────────┬───────────────────────────┤
│ Работа с кодом   │ Работа с файлами │ Анализ и статистика       │
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
│                    Строковые литералы                           │
├─────────────────────────────────────────────────────────────────┤
│ Get: modLiteralsGet*  │  Set: modLiteralsSet*                   │
├─────────────────────────────────────────────────────────────────┤
│                      Классы                                     │
├──────────────────┬──────────────────┬───────────────────────────┤
│ clsAnchors       │ clsObfuscator    │ clsOfficeArchiveManager   │
│ clsLogging       │ clsSort2DArray   │ clsToolsVBACodeStatistics │
│ clsVBECommandHandler │ shSettings   │ ThisWB                    │
└─────────────────────────────────────────────────────────────────┘
```

---

*Справочник сформирован: 03.04.2026*
