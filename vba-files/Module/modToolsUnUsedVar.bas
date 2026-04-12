Attribute VB_Name = "modToolsUnUsedVar"
Option Explicit
Option Private Module

' --- Configuration Constants ---
Private Const COL_OUTPUT_COUNT As Long = 9
Private Const USER_FORM As String = "UserForm"
Private Const CLASS_MODULE As String = "Class Module"

Public Sub showFormUnUsedVariable()
    Dim frmForm     As frmVariableUnUsed
    Set frmForm = New frmVariableUnUsed

    Dim sNameWB     As String
    sNameWB = ActiveWorkbook.Name
    With frmForm
        .Show
        If .lbOK.Caption = "-1" Then
            Set frmForm = Nothing
            Exit Sub
        End If
        With .cmbMain
            If .value = vbNullString Then
                Call MsgBox("No file selected!", vbCritical)
                Exit Sub
            End If
            Dim wb  As Workbook
            Set wb = Workbooks(.value)
            If wb.vbProject.Protection = vbext_pp_locked Then
                Call MsgBox("VBA project is password protected!", vbCritical)
                Exit Sub
            End If
            Dim arr As Variant
            arr = AnalyzeCodeVBProjectUnUsed(wb)
        End With
        Workbooks(sNameWB).Activate
        If IsArrayEmpty(arr) Then
            Debug.Print ">> Analysis completed: all variables are used."
        Else
            .ListCode.List = arr
        End If
        .Show
    End With
End Sub

Private Function IsArrayEmpty(ByRef arr As Variant) As Boolean
    Dim i           As Long
    On Error Resume Next
    i = UBound(arr)
    If Err.Number <> 0 Then IsArrayEmpty = True
    On Error GoTo 0
End Function

' -----------------------------------------------------------------------------
' Main entry point: Analyze VBA project for unused items
' -----------------------------------------------------------------------------
Public Function AnalyzeCodeVBProjectUnUsed(ByRef wb As Workbook) As Variant
    Dim objStats    As clsToolsVBACodeStatistics
    Dim arrModules  As Variant
    Dim arrCodeBase As Variant
    Dim arrDeclarations As Variant
    Dim arrControlsUserForms As Variant
    Dim arrUnusedItems() As String
    Dim objLinkedShapeMacros As Collection

    ' Initialize statistics object
    Set objStats = New clsToolsVBACodeStatistics

    ' 1. Collect project data
    With objStats
        ' Collect procedures
        .addListProcs wb, True
        arrCodeBase = .getArrayCodeBase()
        .reBootArrayCodeBase

        ' Collect modules
        .addListModules wb, True
        arrModules = .getArrayCodeBase()
        .reBootArrayCodeBase

        ' Collect variable declarations
        .addListDeclarations wb, True
        arrDeclarations = .getArrayCodeBase()
        .reBootArrayCodeBase

        ' Collect user form controls
        .addListControlsUserForms wb
        arrControlsUserForms = .getArrayCodeBase()
        .reBootArrayCodeBase
    End With

    ' Check if there is data to analyze
    If IsEmpty(arrModules) Then
        Debug.Print ">> Analysis completed: no data to analyze."
        Exit Function
    End If

    ' Get macros linked to shapes on sheets
    Set objLinkedShapeMacros = GetLinkedShapeMacros(wb)

    ' 2. Find unused items
    arrUnusedItems = FindUnusedItems(arrModules, arrDeclarations, arrCodeBase, arrControlsUserForms, objLinkedShapeMacros, wb)

    AnalyzeCodeVBProjectUnUsed = arrUnusedItems
End Function

' -----------------------------------------------------------------------------
' Logic for finding unused modules, variables, procedures, and parameters
' -----------------------------------------------------------------------------
Private Function FindUnusedItems(ByRef arrModules As Variant, ByRef arrDeclarations As Variant, _
        ByRef arrCodeBase As Variant, ByRef arrControlsUserForms As Variant, _
        ByRef objLinkedShapeMacros As Collection, ByRef wb As Workbook) As String()

    Dim objRegEx    As Object
    Dim objModuleLookup As Collection
    Dim objControlsLookup As Collection
    Dim objClassEventsLookup As Collection
    Dim lUnusedCount As Long
    Dim lMaxItems   As Long

    ' Initialize helper objects
    Set objRegEx = GetRegExObject()
    Set objModuleLookup = GetCollection(arrModules, stdColVBA.stdModuleName)
    Set objControlsLookup = GetControlsLookupCollection(arrControlsUserForms)
    Set objClassEventsLookup = GetClassEventsLookupCollection(arrDeclarations)

    ' Pre-calculate result array size
    lMaxItems = UBound(arrModules, 1)
    If Not IsEmpty(arrDeclarations) Then lMaxItems = lMaxItems + UBound(arrDeclarations, 1)
    If Not IsEmpty(arrCodeBase) Then lMaxItems = lMaxItems + UBound(arrCodeBase, 1)
    If Not IsEmpty(arrControlsUserForms) Then lMaxItems = lMaxItems + UBound(arrControlsUserForms, 1)

    ReDim arrUnused(1 To lMaxItems, 1 To COL_OUTPUT_COUNT) As String

    ' 2.1 Check UserForms and Class Modules
    CheckUnusedModules objRegEx, arrModules, arrUnused, lUnusedCount

    ' 2.2 Check module-level variable declarations
    If Not IsEmpty(arrDeclarations) Then
        CheckUnusedDeclarations objRegEx, arrModules, arrDeclarations, objModuleLookup, arrUnused, lUnusedCount
    End If

    ' 2.3 Check Procedures, Parameters, and Local variables
    If Not IsEmpty(arrCodeBase) Then
        CheckUnusedCodeElements objRegEx, arrModules, arrCodeBase, objModuleLookup, _
                arrUnused, lUnusedCount, objControlsLookup, objClassEventsLookup, objLinkedShapeMacros, wb
    End If

    ' Final trimming of array to actual number of found items
    If lUnusedCount > 0 Then
        Dim arrResult() As String
        Dim i As Long, j As Long

        ReDim arrResult(1 To lUnusedCount, 1 To COL_OUTPUT_COUNT) As String
        For i = 1 To lUnusedCount
            For j = 1 To COL_OUTPUT_COUNT
                arrResult(i, j) = arrUnused(i, j)
            Next j
        Next i
        FindUnusedItems = arrResult
    End If
End Function

' -----------------------------------------------------------------------------
' Find unused modules (Forms and Classes)
' -----------------------------------------------------------------------------
Private Sub CheckUnusedModules(ByRef objRegEx As Object, ByRef arrModules As Variant, _
        ByRef arrUnused() As String, ByRef lUnusedCount As Long)
    Dim i           As Long
    Dim lModulesCount As Long

    lModulesCount = UBound(arrModules, 1)

    For i = 1 To lModulesCount
        Select Case arrModules(i, stdColVBA.stdModuleT)
                 Case USER_FORM, CLASS_MODULE
                ' Check if the module name is used in other modules' code
                If Not FindInAllModulesCode(objRegEx, lModulesCount, arrModules, _
                        arrModules(i, stdColVBA.stdModuleName), _
                        arrModules(i, stdColVBA.stdModuleName), vbNullString) Then
                    AddUnusedItem arrModules, i, arrUnused, lUnusedCount
                End If
        End Select
    Next i
End Sub

' -----------------------------------------------------------------------------
' Find unused module-level variables
' -----------------------------------------------------------------------------
Private Sub CheckUnusedDeclarations(ByRef objRegEx As Object, ByRef arrModules As Variant, _
        ByRef arrDeclarations As Variant, ByRef objModuleLookup As Collection, _
        ByRef arrUnused() As String, ByRef lUnusedCount As Long)

    Dim i           As Long
    Dim lDeclarationsCount As Long
    Dim lModuleIndex As Long
    Dim sCode       As String
    Dim sVarName    As String
    Dim sModifier   As String
    Dim sModuleName As String

    lDeclarationsCount = UBound(arrDeclarations, 1)

    For i = 1 To lDeclarationsCount
        sModifier = arrDeclarations(i, stdColVBA.stdProcModifier)
        sModuleName = arrDeclarations(i, stdColVBA.stdModuleName)
        sVarName = arrDeclarations(i, stdColVBA.stdProcName)
        sVarName = getNameElement(sVarName)

        ' Special case for Enum and Type: name is taken from the definition line
        Select Case arrDeclarations(i, stdColVBA.stdProcType)
                 Case "Enum", "Type"
                sVarName = arrDeclarations(i, stdColVBA.stdProcLines)
        End Select
        lModuleIndex = objModuleLookup(sModuleName)

        ' Exclude the declaration itself from the code being checked (first occurrence only)
        sCode = arrModules(lModuleIndex, stdColVBA.stdCode)
        sCode = VBA.Replace(sCode, arrDeclarations(i, stdColVBA.stdCode), vbNullString, 1, 1)
        sCode = VBA.Replace(sCode, "Dim " & sVarName & " As ", vbNullString)

        Select Case sModifier
                 Case "Private", "Dim"
                ' Variables private to the module. Search for usage only within the module.
                If CountRegexMatches(objRegEx, sCode, sVarName) = 0 Then
                    AddUnusedItem arrDeclarations, i, arrUnused, lUnusedCount
                End If
            Case Else
                ' Public variables, Enum, Type. Search for usage across the entire project.
                If CountRegexMatches(objRegEx, sCode, sVarName) = 0 Then
                    Dim sExcludePattern As String
                    sExcludePattern = " " & sVarName & " As "
                    If Not FindInAllModulesCode(objRegEx, UBound(arrModules, 1), arrModules, _
                            sModuleName, sVarName, sExcludePattern) Then
                        AddUnusedItem arrDeclarations, i, arrUnused, lUnusedCount
                    End If
                End If
        End Select
    Next i
End Sub

' -----------------------------------------------------------------------------
' Find unused procedures, parameters, and local variables
' -----------------------------------------------------------------------------
Private Sub CheckUnusedCodeElements(ByRef objRegEx As Object, ByRef arrModules As Variant, _
        ByRef arrCodeBase As Variant, ByRef objModuleLookup As Collection, _
        ByRef arrUnused() As String, ByRef lUnusedCount As Long, _
        ByRef objControlsLookup As Collection, ByRef objClassEventsLookup As Collection, _
        ByRef objLinkedShapeMacros As Collection, ByRef wb As Workbook)

    Dim i           As Long
    Dim lCodeBaseCount As Long
    Dim lModuleIndex As Long
    Dim sCode       As String
    Dim sProcName   As String
    Dim sProcType   As String
    Dim sElementName As String
    Dim sElementType As String
    Dim sModifier   As String
    Dim sModuleName As String
    Dim sModuleType As String
    Dim bShouldCheck As Boolean
    Dim bLookUI     As Boolean
    Dim objUILookup As Collection

    lCodeBaseCount = UBound(arrCodeBase, 1)

    For i = 1 To lCodeBaseCount
        sModuleName = arrCodeBase(i, stdColVBA.stdModuleName)
        sProcName = arrCodeBase(i, stdColVBA.stdProcName)
        sProcType = arrCodeBase(i, stdColVBA.stdProcType)
        sElementName = arrCodeBase(i, stdColVBA.stdProcLines)
        sElementType = arrCodeBase(i, stdColVBA.stdTypeElement)
        sModifier = arrCodeBase(i, stdColVBA.stdProcModifier)
        sModuleType = arrCodeBase(i, stdColVBA.stdModuleT)

        ' Initialize UI callback lookup (Ribbon) on first encounter of IRibbonControl
        If arrCodeBase(i, stdColVBA.stdProcDeclaration) Like "* As IRibbonControl*" And Not bLookUI Then
            Dim arrUI As Variant
            arrUI = GetArrayFromDictionary(parserLiteralsFormUIOnlyProcedures(wb, True))
            If Not IsEmpty(arrUI) Then
                Set objUILookup = GetCollection(arrUI, 7)
                Erase arrUI
            End If
            bLookUI = True
        End If

        ' Determine whether to check the element (exclude event handlers)
        bShouldCheck = Not IsProcedureEventHandler(sProcName, sModuleType, sModifier, _
                sProcType, objControlsLookup, sModuleName, _
                objClassEventsLookup, objLinkedShapeMacros, objUILookup)

        If bShouldCheck Then
            sElementName = getNameElement(sElementName)
            Select Case sElementType
                     Case "Parametr"
                    ' Ignore IRibbonControl parameters since they are often used implicitly by the system
                    If arrCodeBase(i, stdColVBA.stdProcVariable) <> "IRibbonControl" Then
                        sCode = arrCodeBase(i, stdColVBA.stdCode)
                        If CountRegexMatches(objRegEx, sCode, sElementName) = 0 Then
                            AddUnusedItem arrCodeBase, i, arrUnused, lUnusedCount
                        End If
                    End If

                Case "Local Variable", "Local Const"
                    sCode = arrCodeBase(i, stdColVBA.stdCode)
                    ' Remove the variable/constant declaration line before searching for usage
                    If sElementType = "Local Variable" Then
                        sCode = VBA.Replace(sCode, "Dim " & sElementName & " As ", vbNullString)
                    Else
                        sCode = VBA.Replace(sCode, "Const " & sElementName & " As ", vbNullString)
                    End If

                    If CountRegexMatches(objRegEx, sCode, sElementName) = 0 Then
                        AddUnusedItem arrCodeBase, i, arrUnused, lUnusedCount
                    End If

                Case "Procedure"
                    ' Logic for finding unused procedures
                    If sModifier = "Private" Then
                        ' Private procedure: search for usage only within the current module
                        lModuleIndex = objModuleLookup(sModuleName)
                        sCode = arrModules(lModuleIndex, stdColVBA.stdCode)
                        sCode = VBA.Replace(sCode, arrCodeBase(i, stdColVBA.stdCode), vbNullString, 1, 1)

                        If CountRegexMatches(objRegEx, sCode, sElementName) = 0 Then
                            AddUnusedItem arrCodeBase, i, arrUnused, lUnusedCount
                        End If
                    Else
                        ' Public procedure: search for usage across all modules (except classes, since they have Public API)
                        If sModuleType <> CLASS_MODULE Then
                            lModuleIndex = objModuleLookup(sModuleName)
                            sCode = arrModules(lModuleIndex, stdColVBA.stdCode)
                            sCode = VBA.Replace(sCode, arrCodeBase(i, stdColVBA.stdCode), vbNullString, 1, 1)

                            If CountRegexMatches(objRegEx, sCode, sElementName) = 0 Then
                                If Not FindInAllModulesCode(objRegEx, UBound(arrModules, 1), arrModules, _
                                        sModuleName, sElementName, " " & sElementName & " As ") Then
                                    AddUnusedItem arrCodeBase, i, arrUnused, lUnusedCount
                                End If
                            End If
                        End If
                    End If
            End Select
        End If
    Next i
End Sub

Private Function getNameElement(ByVal sElementName As String) As String
    getNameElement = sElementName
    Dim varType     As String
    varType = typeVariable(getNameElement)
    If VBA.Len(varType) > 0 Then getNameElement = VBA.Left$(getNameElement, VBA.Len(getNameElement) - 1)
End Function

' -----------------------------------------------------------------------------
' Check: is the procedure an event handler (system or UI)
' -----------------------------------------------------------------------------
Private Function IsProcedureEventHandler(ByVal sProcName As String, ByVal sModuleType As String, _
        ByVal sModifier As String, ByVal sProcType As String, _
        ByRef objControlsLookup As Collection, ByVal sModuleName As String, _
        ByRef objClassEventsLookup As Collection, ByRef objLinkedShapeMacros As Collection, _
        ByVal objUILookup As Collection) As Boolean

    ' Exclude Excel auto-procedures
    Select Case sProcName
             Case "Auto_Open", "Auto_Close"
            IsProcedureEventHandler = True
            Exit Function
    End Select

    ' Check references from Ribbon UI
    If haveInCollection(objUILookup, sProcName) > 0 Then
        IsProcedureEventHandler = True
        Exit Function
    End If

    ' Check references from Shapes on sheets
    If haveInCollection(objLinkedShapeMacros, sProcName) > 0 Then
        IsProcedureEventHandler = True
        Exit Function
    End If

    Dim lControlIndex As Long
    IsProcedureEventHandler = False

    ' Analyze only Private Sub
    If sModifier = "Private" And sProcType = "Sub" Then
        Dim i       As Long
        Dim sEventsName As String

        ' WithEvents variable events
        sEventsName = sProcName
        i = VBA.InStrRev(sEventsName, "_") - 1
        If i > 0 Then sEventsName = VBA.Left$(sEventsName, i)

        lControlIndex = haveInCollection(objClassEventsLookup, sModuleName & "." & sEventsName)
        If lControlIndex > 0 Then
            IsProcedureEventHandler = True
            Exit Function
        End If

        ' 1. Workbook and worksheet events (Document Modules)
        If sModuleType = "Document Module" Then
            If sProcName Like "Workbook_*" Or sProcName Like "Worksheet_*" Then
                IsProcedureEventHandler = True
                Exit Function
            End If

            ' 2. UserForm events
        ElseIf sModuleType = USER_FORM Then
            ' Form events (UserForm_Initialize, UserForm_Terminate, etc.)
            If sProcName Like "UserForm_*" Then
                IsProcedureEventHandler = True
                Exit Function
            End If

            ' Control events (CommandButton1_Click, TextBox1_Change, etc.)
            Dim sControlName As String
            sControlName = sProcName
            i = VBA.InStrRev(sControlName, "_") - 1
            If i > 0 Then sControlName = VBA.Left$(sControlName, i)

            lControlIndex = haveInCollection(objControlsLookup, sModuleName & "." & sControlName)
            If lControlIndex > 0 Then
                IsProcedureEventHandler = True
                Exit Function
            End If

            ' 3. Class events
        ElseIf sModuleType = CLASS_MODULE Then
            Select Case sProcName
                     Case "Class_Terminate", "Class_Initialize"
                    IsProcedureEventHandler = True
                    Exit Function
            End Select
        End If
    End If
End Function

' -----------------------------------------------------------------------------
' Helper methods
' -----------------------------------------------------------------------------

' Add a found item to the result array
Private Sub AddUnusedItem(ByRef arrSource As Variant, ByVal lSourceIndex As Long, _
        ByRef arrTarget() As String, ByRef lTargetCount As Long)
    Dim j           As Long
    lTargetCount = lTargetCount + 1
    For j = 1 To COL_OUTPUT_COUNT
        arrTarget(lTargetCount, j) = arrSource(lSourceIndex, j)
    Next j
End Sub

' Safely get a value from a collection by key. Returns 0 if key not found
Private Function haveInCollection(ByRef ObjColl As Collection, ByRef sKey As String) As Long
    If ObjColl Is Nothing Then Exit Function
    On Error Resume Next
    haveInCollection = ObjColl.Item(sKey)
    On Error GoTo 0
End Function

' Create a Lookup collection: Key -> Index in array
Private Function GetCollection(ByRef arr As Variant, ByRef iCol As Integer) As Collection
    Dim objCollection As Collection
    Dim i           As Long
    Dim lCount      As Long

    Set objCollection = New Collection
    lCount = UBound(arr, 1)

    On Error Resume Next    ' Ignore duplicate key errors
    For i = 1 To lCount
        objCollection.Add i, arr(i, iCol)
    Next i
    On Error GoTo 0

    Set GetCollection = objCollection
End Function

' Create a form controls lookup collection: "ModuleName.ControlName" -> Index
Private Function GetControlsLookupCollection(ByRef arrControlsUserForms As Variant) As Collection
    If IsEmpty(arrControlsUserForms) Then
        Set GetControlsLookupCollection = Nothing
        Exit Function
    End If

    Dim objCollection As Collection
    Dim i           As Long
    Dim lCount      As Long

    Set objCollection = New Collection
    lCount = UBound(arrControlsUserForms, 1)

    On Error Resume Next
    For i = 1 To lCount
        objCollection.Add i, arrControlsUserForms(i, stdColVBA.stdModuleName) & "." & arrControlsUserForms(i, stdColVBA.stdProcName)
    Next i
    On Error GoTo 0

    Set GetControlsLookupCollection = objCollection
End Function

' Create a WithEvents variables lookup collection: "ModuleName.VariableName" -> Index
Private Function GetClassEventsLookupCollection(ByRef arrDeclarations As Variant) As Collection
    If IsEmpty(arrDeclarations) Then
        Set GetClassEventsLookupCollection = Nothing
        Exit Function
    End If

    Dim objCollection As Collection
    Dim i           As Long
    Dim lCount      As Long

    Set objCollection = New Collection
    lCount = UBound(arrDeclarations, 1)

    On Error Resume Next
    For i = 1 To lCount
        If arrDeclarations(i, stdColVBA.stdProcType) = "WithEvents" Then
            objCollection.Add i, arrDeclarations(i, stdColVBA.stdModuleName) & "." & arrDeclarations(i, stdColVBA.stdProcName)
        End If
    Next i
    On Error GoTo 0

    Set GetClassEventsLookupCollection = objCollection
End Function

' Search for text in all modules, excluding the current one
Private Function FindInAllModulesCode(ByRef objRegEx As Object, ByRef lCount As Long, ByRef arr As Variant, _
        ByVal sExcludeModuleName As String, ByVal sSearchText As String, _
        ByVal sReplaceText As String) As Boolean
    Dim i           As Long
    Dim sCode       As String

    For i = 1 To lCount
        If arr(i, stdColVBA.stdModuleName) <> sExcludeModuleName Then
            sCode = arr(i, stdColVBA.stdCode)
            If VBA.Len(sReplaceText) > 0 Then sCode = VBA.Replace(sCode, sReplaceText, vbNullString)

            If CountRegexMatches(objRegEx, sCode, sSearchText) > 0 Then
                FindInAllModulesCode = True
                Exit Function
            End If
        End If
    Next i
End Function

' Count the number of regular expression matches
Private Function CountRegexMatches(ByRef objRegEx As Object, ByVal sText As String, ByVal sPattern As String) As Long
    If VBA.Len(sText) = 0 Then Exit Function
    sPattern = RegexEscape(sPattern)
    With objRegEx
        .pattern = "(?:^|[\r\n\(\.\s,=" & Chr$(34) & "])" & sPattern & "(?:$|[\)\.\s,(" & Chr$(34) & "])"
        CountRegexMatches = .Execute(sText).Count
    End With
End Function

' Create and configure the RegExp object
Private Function GetRegExObject() As Object
    Set GetRegExObject = CreateObject("VBScript.RegExp")
    With GetRegExObject
        .Global = True
        .IgnoreCase = True
        .MultiLine = True
    End With
End Function

' Get a collection of macros linked to shapes on sheets
Private Function GetLinkedShapeMacros(ByRef wb As Workbook) As Collection
    Dim sh          As Worksheet
    Dim shp         As Shape
    Dim collMacros  As Collection
    Dim vSplit      As Variant
    Dim strMacroName As String

    Set collMacros = New Collection
    On Error Resume Next

    For Each sh In wb.Worksheets
        For Each shp In sh.Shapes
            If shp.OnAction <> vbNullString Then
                ' Parse OnAction string. Can be "MacroName" or "BookName!MacroName"
                vSplit = VBA.Split(shp.OnAction, "!")

                ' If there is a delimiter, take the second part, otherwise the first (and only)
                If UBound(vSplit) >= 1 Then
                    strMacroName = vSplit(1)
                Else
                    strMacroName = vSplit(0)
                End If

                If Len(strMacroName) > 0 Then
                    collMacros.Add 1, strMacroName
                End If
            End If
        Next shp
    Next sh

    On Error GoTo 0
    Set GetLinkedShapeMacros = collMacros
End Function
