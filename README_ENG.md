**English** | [Русский](README.md)
---

# MACROTools v2.0

> **Powerful Excel VBA Add-in for Developers**
> **Author:** VBATools
> **Version:** 2.0.38
> **License:** Apache License

---

## 📋 Description

**MACROTools** is a professional Excel VBA Add-in that provides an extensive set of tools for VBA project development, analysis, refactoring, and protection.

The add-in integrates into the Visual Basic Editor (VBE) environment via **Ribbon UI** and **context menus**, offering 50+ tools for working with VBA code.

---

## ✨ Features

### 🔧 VBA Code Operations
- **Smart Indenter** — automatic indentation formatting
- **`Dim` Formatting** — merge/split declarations
- **Remove Comments** — clean code from comments
- **Remove Empty Lines** — delete double empty lines
- **Remove Line Continuations** — remove `_` line continuations
- **Swap `=`** — swap left and right sides of assignments
- **Line Numbers** — add/remove line numbers in procedures
- **Debug.Print ON/OFF** — mass enable/disable debug output

### 📊 Statistics and Analysis
- **Module Statistics** — procedure count, line count, types
- **UserForm Statistics** — form controls analysis
- **Declaration Statistics** — variables, constants, types
- **Procedure Statistics** — parameters, modifiers, types
- **Shape Statistics** — macros bound to shapes
- **Unused Variables Search** — dead code detection
- **Unused Modules Search** — unused forms and classes

### 🔐 Protection and Security
- **VBA Password Removal** — bypass VBA project protection
- **Unviewable Protection** — set/remove "Unviewable VBA Project" protection
- **Hide Modules** — hide modules from VBE project window
- **Remove Sheet/Workbook Passwords** — remove protection via XML
- **VBA Obfuscation** — rename variables, procedures, modules

### 📦 File Operations
- **Unpack Office Files** — view internal structure (.xlsx, .xlsm, .xlsb)
- **Pack Office Files** — rebuild archive
- **View Archive Files** — list all files inside archive
- **File Properties** — view/edit built-in and custom document properties

### 🎨 Interface and Themes
- **Dark VBE Theme** — switch to dark theme
- **Light VBE Theme** — switch to light theme
- **Indent Settings** — formatting configuration
- **Comment Settings** — comment templates

### 🔍 Literal Parsing and Renaming
- **String Literal Parsing** — extract strings from code, UserForm, Ribbon UI
- **Literal Renaming** — batch replace string values
- **Character Monitor** — analyze used characters

### 🛠 Utilities
- **MsgBox Constructor** — visual MsgBox generator
- **Formatting Constructor** — format string generator
- **Procedure Constructor** — procedure declaration generator
- **TODO List** — task management in code
- **Code Snippets** — ready-made solutions library
- **Regex Tester** — RegExp debugging
- **Remove External Links** — find and remove external references
- **Toggle A1/R1C1** — Excel reference style
- **Add-in Manager** — manage Excel Add-ins
- **Hotkeys** — hotkey reference

---

## 📂 Project Structure

```
Addin_MacroTools_2.0/
├── vba-files/              # VBA source code
│   ├── Class/              # Class modules (.cls)
│   ├── Form/               # UserForms (.frm)
│   └── Module/             # Standard modules (.bas)
├── Addin_MacroTools_v2.0.38_ENG.xlsb  # Compiled add-in
└── MODULES_REFERENCE.md    # Complete module reference
```

---

## 🚀 Installation

### Manual Installation
1. Copy `Addin_MacroTools_v2.0.38_ENG.xlsb`
2. Open Excel → Click **Install** button

---

### Excel VBE
- Access to VBA object model: **File** → **Options** → **Trust Center** → **Macro Settings** → ✅ Trust access to VBA project object model

---

## ⌨️ Hotkeys

| Combination | Action |
|-------------|--------|
| `Ctrl+Shift+H` | Hotkeys reference |
| `Alt+F11` | Open VBE |

> Full list of hotkeys available via **Tools → Hotkeys** menu

---

## 📖 Documentation

- **[MODULES_REFERENCE.md](MODULES_REFERENCE_ENG.md)** — Complete reference of all modules with procedure descriptions

---

## 🔍 Core Modules

| Module | Description |
|--------|-------------|
| `modAddinConst` | Add-in constants |
| `modAddinCreateMenu` | VBE context menu creation |
| `modAddinPubFun` | Public functions (general) |
| `modAddinPubFunVBE` | Public VBE functions |
| `modAddinPubFunVBEModule` | VBE module operations |
| `modAddinRibbonCallbacks` | Ribbon callback functions |
| `modAddinThemeVBE` | VBE themes |
| `modAddinInstall` | Add-in installation |

### Classes
| Class | Description |
|-------|-------------|
| `clsObfuscator` | VBA project obfuscator |
| `clsOfficeArchiveManager` | Office archive manager |
| `clsToolsVBACodeStatistics` | VBA code statistics |
| `clsLogging` | CSV logger |
| `clsAnchors` | UserForm anchors |
| `clsSort2DArray` | 2D array sorting |

---

## ⚠️ Important Notes

### Security
- Some functions (password removal, obfuscation) use **API hooks** and binary file modification
- VBA protection bypass functions are intended for **restoring access to your own projects**
- Use at your own risk

### VBA Access
- Trusted access to VBA object model is required for proper operation
- Check: `VBAIsTrusted()` in `modAddinPubFunVBE` module

---

## 🐛 Logging

Logs are written to `...\AppData\Roaming\Microsoft\AddIns` folder:
- `MACROTools_logs.csv` — Excel import log

The `clsLogging` class is used to manage logging.

---

## 📝 License

Apache License

---

## 👤 Author

**VBATools**

---

## 🔄 Version

**v2.0.38**

---

## 📞 Support

If you encounter issues:
1. Check log files at `...\AppData\Roaming\Microsoft\AddIns\MACROTools_logs.csv`
2. Ensure VBA access is enabled
3. Restart Excel and verify add-in activation

---

## 🎯 Roadmap

- [ ] Git integration
- [ ] Automated testing
- [ ] API documentation