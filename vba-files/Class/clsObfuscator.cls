VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsObfuscator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Class        :   clsObfuscator - Main class for performing VBA project obfuscation.
'*                                  Encapsulates the logic of parsing, identifier replacement and
'*                                  report generation.
'* Author       :   VBATools
'* Copyright    :   Apache License
'* Created      :   27-03-2026 09:54:06
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

' -- Library references (required in Tools -> References):
' 1. Microsoft Scripting Runtime (Dictionary)
' 2. Microsoft Visual Basic for Applications Extensibility 5.3 (VBProject)

' =============================================================================
' CLASS CONSTANTS AND VARIABLES
' =============================================================================

' -- Configuration --
Private Const mc_REPORT_SHEET As String = "OBFUSCATION_REPORT"
Private Const mc_VARIABLE_SHEET As String = "OBFUSCATION_VARIABLE"
Private Const mc_NAME_MODULE_LITERALS As String = "modLiterals"
Private Const mc_NAME_PREFIX As String = "O"
Private Const mc_NAME_LENGTH As Integer = 12
Private Const mc_MAX_RETRY As Long = 1000
Private Const mc_QUOTE_CHAR As String = """"

' -- Internal state variables --
Private m_wbTarget  As Workbook
Private m_dtStart   As Date

' -- Data collections --
Private m_dictEncoded As Dictionary        ' Dictionary: New name -> Original value
Private m_dictModules As Dictionary        ' Dictionary: Module name -> Module code
Private m_dictControls As Dictionary       ' Dictionary: Control references
Private m_dictMacroShapeMapping As Dictionary    ' Dictionary: Macro -> Shapes
Private m_dictWithEventsClass As Dictionary    ' Dictionary: Classes with WithEvents
Private m_dictUI    As Dictionary

' -- Setting arrays (loaded from sheet) --
Private m_arrModules As Variant
Private m_arrDeclarations As Variant
Private m_arrProcedures As Variant
Private m_arrControls As Variant
Private m_arrProceduresVariable As Variant

' -- RegExp objects --
Private m_objRegExp As Object
Private m_objRegexVar As Object

' =============================================================================
' INITIALIZATION AND TERMINATION
' =============================================================================

Private Sub Class_Initialize()
    Set m_dictEncoded = New Dictionary
    Set m_dictModules = New Dictionary
    ' Initialize RegExp
    Set m_objRegExp = CreateRegExpObject()
    Set m_objRegexVar = CreateRegExpObject()
End Sub

Private Sub Class_Terminate()
    Call ReleaseObjects(m_dictEncoded, m_dictModules, m_dictControls, _
            m_dictMacroShapeMapping, m_dictWithEventsClass, _
            m_objRegExp, m_objRegexVar, m_dictUI)
End Sub

' =============================================================================
' PUBLIC INTERFACE
' =============================================================================

' Main entry point for running obfuscation
Public Function Execute(ByVal wb As Workbook) As Boolean
    'On Error GoTo ErrorHandler

    Set m_wbTarget = wb
    m_dtStart = VBA.Now

    Debug.Print ">> START Obfuscation: " & m_wbTarget.Name

    ' 2. Load settings
    If Not LoadSettings Then Exit Function

    ' 3. Prepare module dictionaries
    Call PrepareModulesDictionary

    ' 4. Backup
    If Not SaveBackupCopy Then Exit Function
    Call LogTime("Backup Created")

    ' 5. Remove old literals module (if exists)
    Call RemoveOldLiteralsModule

    ' -- OBFUSCATION STAGES --

    ' Stage 1: Literals (Strings)
    Call ProcessLiterals
    Call LogTime("Literals Processed")

    ' Stage 2: Local Variables
    If Not IsEmpty(m_arrProceduresVariable) Then
        Call ProcessVariables
        Call LogTime("Variables Processed")
    End If

    ' Stage 3: UserForms Controls
    If Not IsEmpty(m_arrControls) Then
        Call ProcessUserForms
        Call LogTime("UserForms Processed")
    End If

    ' Stage 4: Procedures
    If Not IsEmpty(m_arrProcedures) Then
        Call ProcessProcedures
        Call LogTime("Procedures Processed")
    End If

    ' Stage 5: Declarations
    If Not IsEmpty(m_arrDeclarations) Then
        Call ProcessDeclarations
        Call LogTime("Declarations Processed")
    End If

    ' Stage 6: Modules Renaming
    Call ProcessModulesRename
    Call LogTime("Modules Renamed")

    ' Stage 8: Write changes to project
    Call WriteChangesToProject
    Call LogTime("Changes Written")
    
    ' Stage 7: write UI
    Call writeUI

    ' Stage 9: Generate report
    Call GenerateReport
    Call LogTime("Report Generated")

    ' Finalization
    m_wbTarget.Save
    Execute = True
    Debug.Print ">> FINISH Obfuscation: Success"
    Exit Function

ErrorHandler:
    Debug.Print ">> ERROR in clsObfuscator.Execute: " & Err.Description
    Call WriteErrorLog("clsObfuscator.Execute", False)
    Execute = False
End Function

' =============================================================================
' STAGE LOGIC (STEPS)
' =============================================================================

Private Sub ProcessLiterals()
    Dim arrLiterals As Variant
    arrLiterals = GetArrayFromDictionary(parserLiteralsFormCode(m_wbTarget, True))

    If IsEmpty(arrLiterals) Then Exit Sub

    Dim sCodeLiter  As String
    Dim arrLines()  As String
    Dim i As Long, iCount As Long

    iCount = UBound(arrLiterals, 1)
    ReDim arrLines(1 To iCount) As String

    For i = 1 To iCount
        Dim strModuleName As String
        Dim sVal    As String
        Dim sEncoded As String

        strModuleName = arrLiterals(i, 1)
        sVal = arrLiterals(i, 2)
        sEncoded = AddEncodedName("Literals." & strModuleName & "." & sVal)

        arrLines(i) = "Public Const " & sEncoded & " As String = " & mc_QUOTE_CHAR & sVal & mc_QUOTE_CHAR
        Call ReplaceInModule(strModuleName, mc_QUOTE_CHAR & RegexEscape(sVal) & mc_QUOTE_CHAR & "(?=[\s\(\),]|$)", sEncoded)
    Next i

    ' Create module with constants
    sCodeLiter = "Option Explicit" & vbNewLine & "Option Private Module" & vbNewLine & VBA.Join(arrLines, vbNewLine)

    With m_wbTarget.vbProject.VBComponents.Add(vbext_ct_StdModule)
        .Name = mc_NAME_MODULE_LITERALS
        With .codeModule
            If .CountOfLines > 0 Then Call .DeleteLines(1, .CountOfLines)
            Call .InsertLines(1, sCodeLiter)
        End With
    End With
End Sub

Private Sub ProcessVariables()
    Dim i           As Long
    Dim sProcName   As String
    Dim strPattern  As String
    Dim sTypeElement As String

    For i = 1 To UBound(m_arrProceduresVariable, 1)
        sTypeElement = m_arrProceduresVariable(i, stdColVBA.stdTypeElement)

        Select Case sTypeElement
            Case "Parametr", "Local Variable", "Local Const"
                sProcName = m_arrProceduresVariable(i, stdColVBA.stdProcName)
                strPattern = AddEncodedName(sTypeElement & "." & m_arrProceduresVariable(i, stdColVBA.stdModuleName) & "." & sProcName & "." & m_arrProceduresVariable(i, stdColVBA.stdProcLines))

                Call ReplaceInSpecificSub( _
                        m_arrProceduresVariable(i, stdColVBA.stdModuleName), _
                        m_arrProceduresVariable(i, stdColVBA.stdProcLines), _
                        strPattern, _
                        m_arrProceduresVariable(i, stdColVBA.stdProcType), _
                        sProcName)
        End Select
    Next i
End Sub

Private Sub ProcessUserForms()
    Dim dictControlsNew As Dictionary
    Set dictControlsNew = New Dictionary

    Dim VBComp      As vbComponent
    Dim cnt         As control

    ' Get controls dictionary
    Set m_dictControls = GetDict(m_arrControls, Array(stdColVBA.stdModuleName, stdColVBA.stdProcName), stdColVBA.stdProcName)

    If Not m_dictControls Is Nothing Then
        Dim i As Long, iCount As Long
        Dim sKey As String, sItem As String, arrKeys As Variant
        Dim strEncodedName As String, sNameModule As String
        Dim sCode   As String

        iCount = m_dictControls.Count - 1
        For i = 0 To iCount
            sKey = m_dictControls.keys(i)
            sItem = m_dictControls.Items(i)
            arrKeys = VBA.Split(sKey, ".")
            sNameModule = arrKeys(0)

            Set VBComp = m_wbTarget.vbProject.VBComponents(arrKeys(0))
            Set cnt = VBComp.Designer.Controls(sItem)

            strEncodedName = AddEncodedName("UserForm." & sKey)
            cnt.Name = strEncodedName

            ' Replace in form module code
            sCode = m_dictModules(sNameModule)
            m_dictModules(sNameModule) = ReplaceControlReferences(m_objRegExp, sCode, sItem, strEncodedName)
            ' Replace in all modules
            Call replaceAllModuleControls(sItem, strEncodedName, sNameModule)

            dictControlsNew.Add sNameModule & "." & strEncodedName, strEncodedName
        Next i
    End If

    Set m_dictControls = dictControlsNew
End Sub

Private Sub ProcessProcedures()
    ' Collect macro and shapes mapping
    Set m_dictMacroShapeMapping = GetMacroShapeMapping(m_wbTarget)


    Dim i           As Long
    Dim sProcName As String, sModuleName As String, strPattern As String
    Dim sTypeElement As String

    For i = 1 To UBound(m_arrProcedures, 1)
        sTypeElement = m_arrProcedures(i, stdColVBA.stdTypeElement)

        If sTypeElement = "Procedure" Then
            sProcName = m_arrProcedures(i, stdColVBA.stdProcName)
            sModuleName = m_arrProcedures(i, stdColVBA.stdModuleName)
            strPattern = AddEncodedName(sTypeElement & "." & sModuleName & "." & sProcName)

            If Not IsSystemProc(sProcName, sModuleName) Then
                If m_arrProcedures(i, stdColVBA.stdModuleT) = "UserForm" Then
                    If Not IsUserFormControlEvent(m_arrProcedures(i, stdColVBA.stdProcLines)) Then
                        Call ReplaceInAllModules(sProcName, strPattern, vbNullString)
                    End If
                Else
                    ' Replace procedure name
                    Call ReplaceInAllModules(sProcName, strPattern, vbNullString)

                    ' Handle Ribbon UI
                    If m_arrProcedures(i, stdColVBA.stdProcDeclaration) Like "* As IRibbonControl*" Or _
                            m_arrProcedures(i, stdColVBA.stdProcDeclaration) Like "* As IRibbonUI*" Then
                        If m_dictUI Is Nothing Then Set m_dictUI = New Dictionary
                        Dim arrData(1 To 1, 1 To 2)
                        arrData(1, 1) = sProcName: arrData(1, 2) = strPattern
                        If Not m_dictUI.Exists(sProcName) Then Call m_dictUI.Add(sProcName, arrData)
                    End If

                    ' Handle Shapes
                    If Not m_dictMacroShapeMapping Is Nothing Then
                        If m_dictMacroShapeMapping.Exists(sProcName) Then
                            Call SetNewNameProcInShape(strPattern, m_dictMacroShapeMapping(sProcName))
                        End If
                    End If
                End If
            End If
        End If
    Next i
End Sub

Private Sub ProcessDeclarations()
    Dim i As Long, lCount As Long
    Dim strPattern As String, sTypeElement As String
    Dim strModuleName As String, strWithEventName As String

    lCount = UBound(m_arrDeclarations, 1)

    For i = 1 To lCount
        sTypeElement = m_arrDeclarations(i, stdColVBA.stdTypeElement)

        If sTypeElement = "Declaration" Then
            strModuleName = m_arrDeclarations(i, stdColVBA.stdModuleName)
            strWithEventName = m_arrDeclarations(i, stdColVBA.stdProcLines)
            strPattern = AddEncodedName(sTypeElement & "." & strModuleName & "." & strWithEventName)

            If m_arrDeclarations(i, stdColVBA.stdModuleT) = "Class Module" And _
                    m_arrDeclarations(i, stdColVBA.stdProcType) = "WithEvents" Then

                If Not m_dictWithEventsClass Is Nothing Then
                    If m_dictWithEventsClass.Exists(strModuleName & "." & strWithEventName) Then
                        Call ReplaceInModule(strModuleName, "(Sub\s+)(" & strWithEventName & ")(_[a-zA-Z]*\([^)]*\))", "$1" & strPattern & "$3")
                    End If
                End If
            End If

            Call ReplaceInAllModules(strWithEventName, strPattern, vbNullString)
        End If
    Next i
End Sub

Private Sub ProcessModulesRename()
    Dim arrEncodedNames() As String
    Dim i As Long, lCount As Long

    lCount = m_dictModules.Count
    ReDim arrEncodedNames(0 To lCount - 1) As String

    For i = 0 To lCount - 1
        Dim strModuleName As String
        Dim strEncodedName As String

        strModuleName = m_dictModules.keys(i)
        strEncodedName = AddEncodedName("Module." & strModuleName)
        arrEncodedNames(i) = strEncodedName

        Call ReplaceInAllModules(strModuleName, strEncodedName, vbNullString)
    Next i

    ' Save new names for writing to disk
    ' (Temporary dictionary or property, for simplicity the module-level variable does not store an array of names,
    '  but the WriteChangesToProject method expects that the names in the dictionary are still the old ones.
    '  Therefore here we change the references, and we change the names in VBComponents at the last step)
End Sub

Private Sub WriteChangesToProject()
    Dim i As Long, lCount As Long
    Dim strModuleName As String
    Dim strEncodedName As String

    ' Generate new names for modules (re-use the logic or take from m_dictEncoded)
    ' For simplicity of implementation here we duplicate the generation or use the saved state

    lCount = m_dictModules.Count - 1
    For i = 0 To lCount
        strModuleName = m_dictModules.keys(i)

        ' Generate name (should match what was in ProcessModulesRename)
        ' In real code it is better to save the list of arrEncodedNames in a class property
        ' Here for correctness of the example we search in the encoded dictionary
        Dim sKey    As String
        Dim bFound  As Boolean
        bFound = False

        ' Search for key by value (slow, but works for the example)
        Dim varKey  As Variant
        For Each varKey In m_dictEncoded
            If m_dictEncoded(varKey) = "Module." & strModuleName Then
                strEncodedName = CStr(varKey)
                bFound = True
                Exit For
            End If
        Next varKey

        If bFound Then
            With m_wbTarget.vbProject.VBComponents(strModuleName)
                .Name = strEncodedName
                With .codeModule
                    Dim lCountLines As Long
                    lCountLines = .CountOfLines
                    If lCountLines > 0 Then
                        .DeleteLines 1, lCountLines
                        .InsertLines 1, m_dictModules(strModuleName)
                    End If
                End With
            End With
        End If
    Next i
End Sub

Private Sub GenerateReport()
    With m_dictEncoded
        Dim iCount  As Long
        iCount = .Count
        If iCount = 0 Then Exit Sub

        Dim arrItems As Variant, arrKeys As Variant
        Dim arrRes() As String, arrPart() As String
        Dim i As Long, j As Long

        arrItems = .Items
        arrKeys = .keys
        ReDim arrRes(1 To iCount, 1 To 5)

        For i = 0 To iCount - 1
            arrPart = VBA.Split(arrItems(i), ".", 3)
            For j = 0 To UBound(arrPart)
                arrRes(i + 1, j + 1) = arrPart(j)
            Next j
            arrRes(i + 1, 5) = arrKeys(i)
        Next i

        Call OutputResults(m_wbTarget, mc_REPORT_SHEET, m_wbTarget.FullName, arrRes)
    End With
End Sub

' =============================================================================
' HELPER METHODS
' =============================================================================

' Load settings
Private Function LoadSettings() As Boolean
    Dim sh          As Worksheet
    On Error Resume Next
    Set sh = m_wbTarget.Worksheets(mc_VARIABLE_SHEET)
    If Err.Number <> 0 Then
        MsgBox "Sheet not found [" & mc_VARIABLE_SHEET & "]!", vbCritical
        Exit Function
    End If
    On Error GoTo 0

    Dim lLastRow    As Long
    lLastRow = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
    If lLastRow < 3 Then
        MsgBox "No data on settings sheet!", vbCritical
        Exit Function
    End If

    Dim arrSettings As Variant
    arrSettings = sh.Range(sh.Cells(3, 1), sh.Cells(lLastRow, 11)).Value2

    Call ParseSettingsToArray(arrSettings)
    LoadSettings = True
End Function

Private Sub ParseSettingsToArray(ByRef arrSettings As Variant)
    Dim i As Long, iCount As Long, jCount As Long
    iCount = UBound(arrSettings, 1)
    jCount = UBound(arrSettings, 2)

    Dim arrCounts(1 To 5, 1 To 2) As Long

    ' Counter
    For i = 1 To iCount
        Select Case arrSettings(i, stdColVBA.stdTypeElement)
            Case "Module": arrCounts(1, 1) = arrCounts(1, 1) + 1
            Case "Declaration": arrCounts(2, 1) = arrCounts(2, 1) + 1
            Case "Procedure": arrCounts(3, 1) = arrCounts(3, 1) + 1
            Case "UserForms": arrCounts(4, 1) = arrCounts(4, 1) + 1
            Case "Parametr", "Local Variable", "Local Const": arrCounts(5, 1) = arrCounts(5, 1) + 1
        End Select
    Next i

    ' ReDim
    If arrCounts(1, 1) > 0 Then ReDim m_arrModules(1 To arrCounts(1, 1), 1 To jCount)
    If arrCounts(2, 1) > 0 Then ReDim m_arrDeclarations(1 To arrCounts(2, 1), 1 To jCount)
    If arrCounts(3, 1) > 0 Then ReDim m_arrProcedures(1 To arrCounts(3, 1), 1 To jCount)
    If arrCounts(4, 1) > 0 Then ReDim m_arrControls(1 To arrCounts(4, 1), 1 To jCount)
    If arrCounts(5, 1) > 0 Then ReDim m_arrProceduresVariable(1 To arrCounts(5, 1), 1 To jCount)

    ' Fill
    For i = 1 To iCount
        Select Case arrSettings(i, stdColVBA.stdTypeElement)
            Case "Module"
                arrCounts(1, 2) = arrCounts(1, 2) + 1
                arrSettings(i, stdColVBA.stdCode) = clearCodeStrings(GetCodeFromModule(m_wbTarget.vbProject.VBComponents(arrSettings(i, stdColVBA.stdModuleName)).codeModule))
                Call CopyRow(arrSettings, m_arrModules, i, arrCounts(1, 2), jCount)
            Case "Declaration"
                arrCounts(2, 2) = arrCounts(2, 2) + 1
                Call CopyRow(arrSettings, m_arrDeclarations, i, arrCounts(2, 2), jCount)
            Case "Procedure"
                arrCounts(3, 2) = arrCounts(3, 2) + 1
                Call CopyRow(arrSettings, m_arrProcedures, i, arrCounts(3, 2), jCount)
            Case "UserForms"
                arrCounts(4, 2) = arrCounts(4, 2) + 1
                Call CopyRow(arrSettings, m_arrControls, i, arrCounts(4, 2), jCount)
            Case "Parametr", "Local Variable", "Local Const"
                arrCounts(5, 2) = arrCounts(5, 2) + 1
                Call CopyRow(arrSettings, m_arrProceduresVariable, i, arrCounts(5, 2), jCount)
        End Select
    Next i

    Set m_dictWithEventsClass = GetWithEventsClassDict(m_arrDeclarations)
End Sub

Private Function GetWithEventsClassDict(ByRef arrDeclaration As Variant) As Dictionary

    If IsEmpty(arrDeclaration) Then Exit Function

    Dim dictWithEventsClass As Dictionary
    Set dictWithEventsClass = New Dictionary
    Dim i           As Long
    Dim iCount      As Long
    Dim sKey        As String
    iCount = UBound(arrDeclaration)
    With dictWithEventsClass
        For i = 1 To iCount
            If arrDeclaration(i, stdColVBA.stdProcType) = "WithEvents" Then
                sKey = arrDeclaration(i, stdColVBA.stdModuleName) & "." & arrDeclaration(i, stdColVBA.stdProcName)
                If Not .Exists(sKey) Then
                    .Add sKey, arrDeclaration(i, stdColVBA.stdProcName)
                End If
            End If
        Next i
    End With
    Set GetWithEventsClassDict = dictWithEventsClass
End Function

Private Sub CopyRow(ByRef arrSource As Variant, ByRef arrDest As Variant, _
        ByVal lSourceRow As Long, ByVal lDestRow As Long, _
        ByVal lCols As Long)
    Dim j           As Long
    For j = 1 To lCols
        arrDest(lDestRow, j) = arrSource(lSourceRow, j)
    Next j
End Sub

Private Sub PrepareModulesDictionary()
    If IsEmpty(m_arrModules) Then Exit Sub
    Set m_dictModules = GetDict(m_arrModules, stdColVBA.stdModuleName, stdColVBA.stdCode)
End Sub

Private Sub RemoveOldLiteralsModule()
    On Error Resume Next
    With m_wbTarget.vbProject
        Call .VBComponents.Remove(.VBComponents(mc_NAME_MODULE_LITERALS))
    End With
    On Error GoTo 0
End Sub

' -- Replacements --

Private Sub replaceAllModuleControls(ByVal sOldName As String, ByVal sNewName As String, ByVal sExceptModule As String)
    Dim varKey      As Variant
    Dim sCode       As String
    Dim hasChanges  As Boolean
    Dim sNameModule As String

    For Each varKey In m_dictModules.keys
        sNameModule = CStr(varKey)
        If sNameModule <> sExceptModule Then
            sCode = m_dictModules(sNameModule)
            If sCode <> vbNullString Then
                sCode = ReplaceControlName(m_objRegExp, sCode, sExceptModule, sOldName, sNewName, hasChanges)
                If hasChanges Then
                    m_dictModules(sNameModule) = sCode
                End If
            End If
        End If
    Next varKey
End Sub

Private Sub ReplaceInAllModules(ByVal sOldName As String, ByVal sNewName As String, ByVal sExceptModule As String)
    Dim varKey      As Variant
    For Each varKey In m_dictModules.keys
        If CStr(varKey) <> sExceptModule Then
            Call ReplaceInModule(CStr(varKey), GetWordBoundaryPattern(sOldName), "$1" & sNewName & "$2")
        End If
    Next varKey
End Sub

Private Sub ReplaceInModule(ByVal sModuleName As String, ByVal sPattern As String, ByVal sReplacement As String)
    If m_dictModules.Exists(sModuleName) Then
        Dim sCode   As String
        sCode = m_dictModules(sModuleName)
        If Len(sCode) > 0 Then
            sCode = ExecuteRegExpReplace(m_objRegExp, sCode, sPattern, sReplacement)
            m_dictModules(sModuleName) = sCode
        End If
    End If
End Sub

Private Sub ReplaceInSpecificSub(ByVal sModuleName As String, ByVal strNameReplace As String, ByVal strPattern As String, ByVal sTypeProc As String, ByVal sNameProc As String)
    Dim sCode       As String
    sCode = m_dictModules(sModuleName)

    ' Regex to extract the procedure block
    m_objRegExp.pattern = "(" & sTypeProc & " " & sNameProc & "\((?:""(?:[^""""]|"""")*""|'[^\r\n]*|[\s\S])*?End " & sTypeProc & "\b)"

    If m_objRegExp.test(sCode) Then
        Dim colMatches As Object, matchBlock As Object
        Set colMatches = m_objRegExp.Execute(sCode)
        Set matchBlock = colMatches(0)

        Dim strBlockText As String
        strBlockText = matchBlock.value

        m_objRegexVar.pattern = GetWordBoundaryPattern(strNameReplace)
        Dim strNewBlockText As String
        strNewBlockText = m_objRegexVar.Replace(strBlockText, "$1" & strPattern & "$2")

        m_dictModules(sModuleName) = Left(sCode, matchBlock.FirstIndex) & strNewBlockText & mid(sCode, matchBlock.FirstIndex + matchBlock.Length + 1)
    End If
End Sub

' -- Checks --

Private Function IsSystemProc(ByRef sProcName As String, ByRef sModuleName As String) As Boolean
    Select Case sProcName
        Case "Class_Terminate", "Class_Initialize", "Auto_Close", "Auto_Open": IsSystemProc = True: Exit Function
    End Select

    Dim arr         As Variant
    arr = Split(sProcName, "_")
    If UBound(arr) = 1 Then
        Select Case arr(0)
            Case "Workbook", "Worksheet", "UserForm": IsSystemProc = True: Exit Function
        End Select
        If Not m_dictWithEventsClass Is Nothing Then
            If m_dictWithEventsClass.Exists(sModuleName & "." & arr(0)) Then IsSystemProc = True: Exit Function
        End If
    End If
End Function

Private Function IsUserFormControlEvent(ByVal strProcLines As String) As Boolean
    If m_dictControls Is Nothing Then Exit Function
    Dim iPos        As Integer
    iPos = VBA.InStrRev(strProcLines, "_") - 1
    If iPos > 0 Then
        IsUserFormControlEvent = m_dictControls.Exists(VBA.Left$(strProcLines, iPos))
    End If
End Function

' -- Utilities --

Private Function AddEncodedName(ByVal sValue As String) As String
    Dim strKey As String, lRetry As Long

    Do
        strKey = mc_NAME_PREFIX & GenerateBinaryString(mc_NAME_LENGTH)
        lRetry = lRetry + 1
        If lRetry > mc_MAX_RETRY Then Err.Raise vbObjectError + 1001, , "Max retry exceeded"
    Loop While m_dictEncoded.Exists(strKey)

    m_dictEncoded.Add strKey, sValue
    AddEncodedName = strKey
End Function

Private Function GenerateBinaryString(ByVal lLength As Integer) As String
    Dim arrChars() As Byte, i As Long
    ReDim arrChars(0 To lLength - 1)
    Dim arrValid(0 To 3) As Byte
    arrValid(0) = 79: arrValid(1) = 48: arrValid(2) = 49: arrValid(3) = 108    ' O, 0, 1, l

    For i = 0 To lLength - 1
        arrChars(i) = arrValid(Int(Rnd() * 4))
    Next i
    GenerateBinaryString = VBA.StrConv(arrChars, vbUnicode)
End Function

Private Function CreateRegExpObject() As Object
    Dim obj         As Object
    Set obj = CreateObject("VBScript.RegExp")
    With obj
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
    End With
    Set CreateRegExpObject = obj
End Function

Private Function ExecuteRegExpReplace(ByRef obj As Object, ByVal sText As String, ByVal sPattern As String, ByVal sRepl As String) As String
    If Len(sText) = 0 Then ExecuteRegExpReplace = sText: Exit Function
    obj.pattern = sPattern
    ExecuteRegExpReplace = obj.Replace(sText, sRepl)
End Function

Private Function GetWordBoundaryPattern(ByVal sName As String) As String
    GetWordBoundaryPattern = "(^|[^a-zA-Z0-9_\" & Chr(34) & "])" & sName & "(?!:=)([^a-zA-Z0-9_\" & Chr(34) & "]|$)"
End Function

Private Function RegexEscape(ByVal sInput As String) As String
    Dim sSpecial As String, i As Long, sChar As String
    sSpecial = "\^$.|?*+()[]{}"
    For i = 1 To Len(sSpecial)
        sChar = mid(sSpecial, i, 1)
        sInput = Replace(sInput, sChar, "\" & sChar)
    Next i
    RegexEscape = sInput
End Function

Private Function SaveBackupCopy() As Boolean
    Dim sWBName     As String
    Application.DisplayAlerts = False
    sWBName = sGetBaseName(m_wbTarget.Name) & "_obf." & sGetExtensionName(m_wbTarget.Name)
    If WorkbookIsOpen(sWBName) Then
        MsgBox "Close the workbook " & sWBName, vbCritical
    Else
        m_wbTarget.SaveAs m_wbTarget.Path & Application.PathSeparator & sWBName
        SaveBackupCopy = True
    End If
    Application.DisplayAlerts = True
End Function

Private Sub LogTime(ByVal sStep As String)
    Debug.Print ">> " & VBA.Format$(VBA.Now() - m_dtStart, "hh:mm:ss") & " " & sStep
End Sub

Private Sub ReleaseObjects(ParamArray arrObjects() As Variant)
    Dim vObj        As Variant
    On Error Resume Next
    For Each vObj In arrObjects
        If Not vObj Is Nothing Then
            If TypeOf vObj Is Object  Then Set vObj = Nothing
            ' If it's a Dictionary, also release (although in VBA Set Nothing is enough for objects)
        End If
    Next vObj
    On Error GoTo 0
End Sub
Private Function GetDict(ByRef arr As Variant, ByVal arrColKey As Variant, ByVal lColItem As Long) As Dictionary
    Dim objDict     As Dictionary
    Dim i As Long, lCount As Long

    Set objDict = New Dictionary
    On Error Resume Next
    lCount = UBound(arr, 1)
    If Err.Number <> 0 Then Exit Function
    On Error GoTo 0

    If IsNumeric(arrColKey) Then
        arrColKey = Array(arrColKey)
    End If

    Dim sKey        As String
    Dim j           As Integer

    With objDict
        For i = 1 To lCount
            sKey = vbNullString
            For j = 0 To UBound(arrColKey, 1)
                If sKey <> vbNullString Then sKey = sKey & "."
                sKey = sKey & arr(i, arrColKey(j))
            Next j
            If Not .Exists(sKey) Then
                .Add sKey, arr(i, lColItem)
            End If
        Next i
    End With

    Set GetDict = objDict
End Function

Private Function GetMacroShapeMapping(ByRef wb As Workbook) As Dictionary
    Dim sh          As Worksheet
    Dim shp         As Shape
    Dim sProc       As String
    Dim iPos        As Integer
    Dim oDic        As Dictionary

    For Each sh In wb.Worksheets
        For Each shp In sh.Shapes
            sProc = shp.OnAction
            If sProc <> vbNullString Then
                If oDic Is Nothing Then
                    Set oDic = New Dictionary
                    oDic.CompareMode = vbTextCompare
                End If
                iPos = VBA.InStrRev(sProc, "!")
                If iPos > 0 Then
                    sProc = VBA.Right$(sProc, VBA.Len(sProc) - iPos)
                End If
                If Not oDic.Exists(sProc) Then
                    oDic.Add sProc, sh.Name & "}.{" & shp.Name
                Else
                    oDic(sProc) = oDic(sProc) & "}|{" & sh.Name & "}.{" & shp.Name
                End If
            End If
        Next shp
    Next sh
    Set GetMacroShapeMapping = oDic
End Function

' =============================================================================
' UI REFERENCES UPDATE (Ribbon)
' =============================================================================

Private Sub writeUI()
    If m_dictUI Is Nothing Then Exit Sub
    Dim dictHookUI  As Dictionary
    Dim sFullName   As String
    Dim sFolderName As String
    Const PREFIX_UNZIPPED As String = "UnZipped_"

    sFullName = m_wbTarget.FullName
    Set dictHookUI = parserLiteralsFormUIOnlyProcedures(m_wbTarget, False)
    If dictHookUI Is Nothing Then Exit Sub

    Dim arrHookUI   As Variant
    Dim arrUI       As Variant


    arrHookUI = GetArrayFromDictionary(dictHookUI)
    arrUI = GetArrayFromDictionary(m_dictUI)

    Dim i           As Long
    Dim iCount      As Long
    Dim j           As Long
    Dim jCount      As Long
    Dim oXMLDoc     As MSXML2.DOMDocument
    Dim sCurrentKey As String
    Dim sPathUI     As String

    iCount = UBound(arrHookUI, 1)
    jCount = UBound(arrUI, 1)

    Dim mFSO        As Object
    Set mFSO = CreateObject("Scripting.FileSystemObject")

    sFullName = VBA.Replace(sFullName, "—", "-")
    With mFSO
        sFolderName = .GetParentFolderName(sFullName) & Application.PathSeparator & PREFIX_UNZIPPED & .getFileName(sFullName)
    End With
    Set mFSO = Nothing

    For i = 1 To iCount
        For j = 1 To jCount
            If arrHookUI(i, UIColumns.UI_AttrText) = arrUI(j, 1) Then
                arrHookUI(i, UIColumns.UI_AttrTextNew) = arrUI(j, 2)
                If sCurrentKey <> CStr(arrHookUI(i, UIColumns.UI_ModuleType)) Then
                    ' Save the previous document if it was loaded
                    If Not oXMLDoc Is Nothing Then
                        Call oXMLDoc.Save(sPathUI)
                    End If

                    ' Setup for the new type
                    sCurrentKey = CStr(arrHookUI(i, UIColumns.UI_ModuleType))
                    Select Case sCurrentKey
                        Case "CustomUI"
                            sPathUI = sFolderName & Application.PathSeparator & "customUI" & Application.PathSeparator & "customUI.xml"
                        Case "CustomUI14"
                            sPathUI = sFolderName & Application.PathSeparator & "customUI" & Application.PathSeparator & "customUI14.xml"
                        Case Else
                            sPathUI = vbNullString
                    End Select
                    If FileHave(sPathUI, vbNormal) Then Set oXMLDoc = getXMLDOC(sPathUI)
                End If
                Call WriteXML(arrHookUI, i, oXMLDoc)
                Exit For
            End If
        Next j
        If Not oXMLDoc Is Nothing Then Call oXMLDoc.Save(sPathUI)
    Next i
    
    Call ZipAllFilesInFolder(CreateObject("Scripting.FileSystemObject"), sFolderName, sFullName)
    
    Set m_wbTarget = Workbooks.Open(sFullName)
    Call LogTime("Write UI Ribbon")
End Sub

Private Function getXMLDOC(ByVal sPathXML As String) As MSXML2.DOMDocument
    ' Check if file exists
    If Not FileHave(sPathXML, vbNormal) Then Exit Function

    Dim oXMLDoc     As MSXML2.DOMDocument
    Set oXMLDoc = New MSXML2.DOMDocument
    With oXMLDoc
        .async = False
        .validateOnParse = False
        If Not .Load(sPathXML) Then Exit Function
        If .parseError.ErrorCode <> 0 Then Exit Function
        Set getXMLDOC = oXMLDoc
    End With
End Function


'    ' Finalize UI
'    With dictPoceduresNameFormUI
'        If Not .dictUIName Is Nothing Then
'            Dim arrVal As Variant
'            Dim iCount As Long
'            Dim j   As Long
'            arrVal = .dictUIName.Items()
'            iCount = .dictUIName.Count
'            ReDim arrUI(1 To iCount, UIColumns.[_First] To UIColumns.[_Last])
'            Set .dictProcName = Nothing
'            Set .dictUIName = Nothing
'            For i = 1 To iCount
'                For j = UIColumns.[_First] To UIColumns.[_Last]
'                    arrUI(i, j) = arrVal(i - 1)(1, j)
'                Next j
'            Next i
'            Call renameLiteralsToUI(m_wbTarget, arrUI)
'            Debug.Print "    >> Rename UI link Procedure"
'        End If
'    End With

' =============================================================================
' SHAPE REFERENCES UPDATE (SHAPES)
' =============================================================================

' Updates the macro assigned to shapes (buttons) on Excel worksheets
' sShapeMappingData - mapping string formed by the GetMacroShapeMapping function
'                     Format: "Sheet1}.{Shape1}|{Sheet2}.{Shape2"
Private Sub SetNewNameProcInShape(ByVal sProcNameEncoded As String, _
        ByVal sShapeMappingData As String)

    If VBA.Len(sShapeMappingData) = 0 Then Exit Sub

    Dim i           As Long
    Dim iCount      As Long
    Dim arrShapes   As Variant
    Dim arrShapeDetails As Variant
    Dim sSheetName  As String
    Dim sShapeName  As String
    Dim sNewOnAction As String

    arrShapes = VBA.Split(sShapeMappingData, "}|{")
    iCount = UBound(arrShapes)

    For i = 0 To iCount
        arrShapeDetails = VBA.Split(arrShapes(i), "}.{")

        If UBound(arrShapeDetails) >= 1 Then
            sSheetName = arrShapeDetails(0)
            sShapeName = arrShapeDetails(1)

            sNewOnAction = "'" & m_wbTarget.Name & "'!" & sProcNameEncoded

            If Err.Number = 0 Then
                m_wbTarget.Worksheets(sSheetName).Shapes(sShapeName).OnAction = sNewOnAction
            End If

            If Err.Number <> 0 Then
                Debug.Print ">> Warning: Failed to update shape [" & sSheetName & "]." & sShapeName
                Err.Clear
            End If
        End If
    Next i
End Sub