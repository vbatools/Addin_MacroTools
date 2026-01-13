# Обзор проекта MACROTools

## Название проекта
MACROTools - это надстройка для Microsoft Excel, предназначенная для автоматизации и улучшения процесса разработки VBA кода.

## Назначение
Проект представляет собой надстройку Excel (.xlam), которая добавляет в Visual Basic Editor (VBE) различные инструменты для упрощения написания, форматирования и управления VBA кодом.

## Архитектура проекта

### Структура проекта
```
vba-files/
├── Class/              # Классы VBA
│   ├── clsVBECommandHandler.cls  # Обработчик команд VBE
│   ├── shSettings.cls            # Лист настроек
│   └── ThisWB.cls                # Класс книги
├── Form/               # Формы VBA
│   ├── frmOptionsModule.frm      # Форма опций модуля
│   └── frmSettingsIndent.frm     # Форма настроек отступов
└── Module/             # Модули VBA
    ├── modAddinConst.bas         # Константы аддина
    ├── modAddinCreateMenus.bas   # Создание меню
    ├── modAddinInstall.bas       # Установка аддина
    ├── modAddinPubFun.bas        # Публичные функции
    ├── modAddinPubFunVBE.bas     # Публичные функции VBE
    ├── modAddinRibbonCallbacks.bas # Callback-функции ленты
    ├── modToolsDebugOnOff.bas    # Включение/выключение отладки
    ├── modToolsDelTwoEmptyStrings.bas # Удаление пустых строк
    ├── modToolsDimOneLine.bas    # Работа с объявлениями переменных
    ├── modToolsLineIndent.bas    # Инструменты отступов
    ├── modToolsLineNumbers.bas   # Работа с номерами строк
    ├── modToolsOptionsModule.bas # Опции модуля
    ├── modToolsOther.bas         # Прочие инструменты
    ├── modToolsSwapEgual.bas     # Замена знаков равенства
    ├── modUFControlsAlingHorizVert.bas # Выравнивание контролов
    ├── modUFControlsLowerUpperCase.bas # Преобразование регистра
    ├── modUFControlsMove.bas     # Перемещение контролов
    ├── modUFControlsReName.bas   # Переименование контролов
    └── modUFControlsStyleCopyPaste.bas # Копирование/вставка стилей
```

## Основные функциональности

### 1. Управление меню и интерфейсом
- `modAddinCreateMenus.bas` - создает контекстные меню в VBE
- `clsVBECommandHandler.cls` - обрабатывает нажатия кнопок в VBE

### 2. Форматирование кода
- `modToolsLineIndent.bas` - продвинутый инструмент форматирования отступов (основанный на Smart Indenter Стивена Буллена)
- Настройки форматирования доступны через форму `frmSettingsIndent.frm`

### 3. Управление отладкой
- `modToolsDebugOnOff.bas` - включение/выключение отладочных сообщений (комментирование/раскомментирование строк Debug.Print)

### 4. Работа с формами и контролами
- `modUFControlsReName.bas` - переименование контролов с обновлением кода
- `modUFControlsAlingHorizVert.bas` - выравнивание контролов
- `modUFControlsStyleCopyPaste.bas` - копирование/вставка стилей контролов
- `modUFControlsMove.bas` - перемещение контролов

### 5. Управление опциями VBA
- `modToolsOptionsModule.bas` - добавление опций (Option Explicit, Option Private Module и т.д.)
- Форма `frmOptionsModule.frm` - графический интерфейс для выбора опций

### 6. Работа с объявлением переменных
- `modToolsDimOneLine.bas` - преобразование многострочных объявлений в однострочные
- `modToolsOther.bas` - другие инструменты для работы с переменными

### 7. Работа с номерами строк
- `modToolsLineNumbers.bas` - добавление/удаление номеров строк

## Технические особенности

### Установка и интеграция
- `modAddinInstall.bas` - процедура установки надстройки
- Автоматическое добавление меню в VBE при запуске
- Поддержка контекстных меню для различных частей среды VBE

### Поддержка настроек
- Использование листа Settings для хранения настроек
- Формы для удобной настройки параметров

### Безопасность и надежность
- Обработка ошибок во всех основных процедурах
- Поддержка отката изменений (Undo) для некоторых операций

## Пользовательские меню

Аддин добавляет следующие меню в VBE:

1. **MENU_MOVE_CONTROLS** - инструменты для перемещения контролов
2. **MENU_TOOLS** - основные инструменты для работы с кодом
3. **MENU_CODE_WINDOW** - инструменты для работы с окном кода
4. **MENU_FORMS** - инструменты для работы с формами
5. **MENU_PROJECT_WINDOW** - инструменты для работы с проектом
6. **MENU_MS_FORMS** - инструменты для работы с MS Forms

## Заключение

MACROTools представляет собой комплексный инструмент для повышения производительности разработчиков VBA, обеспечивая широкий спектр возможностей для форматирования, рефакторинга и управления VBA кодом в среде Excel.