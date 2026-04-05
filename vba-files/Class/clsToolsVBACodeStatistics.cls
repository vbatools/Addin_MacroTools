VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsToolsVBACodeStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================
' CLASS: CodeParser
' PURPOSE: Analyze VBA project structure (modules, procedures, declarations)
' ======================================================

' --- Constants for JSON keys ---

Private Const JSON_KEY_MODULES As String = "modules"
Private Const JSON_KEY_MODULE_NAME As String = "moduleName"
Private Const JSON_KEY_MODULE_TYPE As String = "moduleType"
Private Const JSON_KEY_GLOBAL_DECLARATIONS As String = "globalDeclarations"
Private Const JSON_KEY_PROCEDURES As String = "procedures"
Private Const JSON_KEY_CONTROLS As String = "controls"
Private Const JSON_KEY_NAME As String = "name"
Private Const JSON_KEY_MODIFIER As String = "modifier"
Private Const JSON_KEY_DATA_TYPE As String = "dataType"
Private Const JSON_KEY_VARIABLE_TYPE As String = "variableType"
Private Const JSON_KEY_VALUE As String = "value"
Private Const JSON_KEY_CONTROL_TYPE As String = "type"
Private Const JSON_KEY_CONTROL_TIP_TEXT As String = "controlTipText"
Private Const JSON_KEY_PARAMETRS As String = "parameters"
Private Const JSON_KEY_LOCAL_VARIABLES As String = "localVariables"
Private Const JSON_KEY_LOCAL_CONSTS As String = "localConsts"
Private Const JSON_KEY_WORKBOOK_NAME As String = "workbookName"
Private Const JSON_KEY_WORKBOOK_FULL_NAME As String = "workbookFullName"
Private Const JSON_KEY_VB_PROJECT_NAME As String = "vbProjectName"
Private Const JSON_KEY_WORKBOOK_AUTHOR As String = "workbookAuthor"
Private Const JSON_KEY_WORKBOOK_LAST_AUTHOR As String = "workbookLastAuthor"
Private Const JSON_KEY_WORKBOOK_CREATION_DATE As String = "workbookCreationDate"
Private Const JSON_KEY_WORKBOOK_LAST_SAVE_DATE As String = "workbookLastSaveDate"
Private Const JSON_KEY_COUNT_MODULES As String = "countModules"
Private Const COMMA As String = ", "

' --- Enumerations and Constants ---

Public Enum stdColVBA
    stdTypeElement = 1
    stdModuleT
    stdModuleName
    stdProcModifier
    stdProcType
    stdProcName
    stdProcLines
    stdProcVariable
    stdProcDeclaration
    stdProcComment
    stdCode
    [_First] = stdModuleT
    [_Last] = stdCode
End Enum

Private Const MAX_ITERATION_LIMIT As Long = 10000
Private Const COMMENT_SCAN_LINES As Long = 30
Private Const COLUMN_COUNT As Byte = stdColVBA.[_Last]

' --- Class variables ---
Private mVbProj     As VBIDE.vbProject
Private mWB         As Workbook
Private mDic        As Dictionary

' ======================================================
' Public methods: Data collection
' ======================================================
Public Sub addListModules(ByRef wb As Workbook, ByRef getCode As Boolean)
    Call InitializeContext(wb)
    If mVbProj Is Nothing Then Exit Sub

    Dim oComp       As VBIDE.vbComponent
    Dim sModuleKey  As String
    Dim arrData()   As String

    For Each oComp In mVbProj.VBComponents
        With oComp
            sModuleKey = moduleTypeName(.Type) & "." & .Name
            If Not mDic.Exists(sModuleKey) Then
                ' IMPORTANT: ReDim inside the loop is necessary to create a new array for each element.
                ' Otherwise, the dictionary will store references to the same array object.
                ReDim arrData(1 To 1, 1 To COLUMN_COUNT) As String
                arrData(1, stdColVBA.stdTypeElement) = "Module"
                arrData(1, stdColVBA.stdModuleT) = moduleTypeName(.Type)
                arrData(1, stdColVBA.stdModuleName) = .Name
                With .codeModule
                    If .CountOfLines > 1 And getCode Then arrData(1, stdColVBA.stdCode) = clearCodeStrings(.Lines(1, .CountOfLines))
                End With
                mDic.Add sModuleKey, arrData
            End If
        End With
    Next oComp
End Sub

Public Sub addListProcs(ByRef wb As Workbook, ByRef getCode As Boolean)
    Call InitializeContext(wb)
    If mVbProj Is Nothing Then Exit Sub

    Dim oComp       As VBIDE.vbComponent
    For Each oComp In mVbProj.VBComponents
        Call ParserProceduresInModule(oComp.codeModule, getCode)
    Next oComp
End Sub

Public Sub addListDeclarations(ByRef wb As Workbook, ByRef getCode As Boolean)
    Call InitializeContext(wb)
    If mVbProj Is Nothing Then Exit Sub

    Dim oComp       As VBIDE.vbComponent
    For Each oComp In mVbProj.VBComponents
        Call ParseDeclarationSection(oComp.codeModule, getCode)
    Next oComp
End Sub

Public Sub addListControlsUserForms(ByRef wb As Workbook)
    Call InitializeContext(wb)
    If mVbProj Is Nothing Then Exit Sub

    Dim oComp       As VBIDE.vbComponent
    For Each oComp In mVbProj.VBComponents
        Call ParserControlsUserForm(oComp.codeModule)
    Next oComp
End Sub

' ======================================================
' Public methods: Data export
' ======================================================

Public Sub reBootArrayCodeBase()
    If Not mDic Is Nothing Then mDic.RemoveAll
End Sub

Public Function getArrayCodeBase() As Variant
    Dim arr         As Variant
    arr = GetArrayFromDictionary(mDic)
    If IsEmpty(arr) Then Exit Function

    Dim i           As Long
    Dim iCount      As Long
    iCount = UBound(arr, 1)
    ReDim arrFilter(1 To iCount, 1 To 2)
    For i = 1 To iCount
        arrFilter(i, 1) = arr(i, stdColVBA.stdModuleName) & "|"

        Select Case arr(i, stdColVBA.stdTypeElement)
                 Case "Module"
                arrFilter(i, 1) = arrFilter(i, 1) & 0 & "|" & 0 & "|" & 0 & "|" & 0
            Case "Declaration"
                arrFilter(i, 1) = arrFilter(i, 1) & 1 & "|" & arr(i, stdColVBA.stdProcName) & "|" & 0 & "|" & 0
            Case "UserForms"
                arrFilter(i, 1) = arrFilter(i, 1) & 2 & "|" & arr(i, stdColVBA.stdProcName) & "|" & 0 & "|" & 0
            Case "Procedure"
                arrFilter(i, 1) = arrFilter(i, 1) & 3 & "|" & arr(i, stdColVBA.stdProcName) & "|" & 0 & "|" & 0
            Case "Parametr"
                arrFilter(i, 1) = arrFilter(i, 1) & 3 & "|" & arr(i, stdColVBA.stdProcName) & "|" & 1 & "|" & arr(i, stdColVBA.stdProcLines)
            Case "Local Const"
                arrFilter(i, 1) = arrFilter(i, 1) & 3 & "|" & arr(i, stdColVBA.stdProcName) & "|" & 2 & "|" & arr(i, stdColVBA.stdProcLines)
            Case "Local Variable"
                arrFilter(i, 1) = arrFilter(i, 1) & 3 & "|" & arr(i, stdColVBA.stdProcName) & "|" & 3 & "|" & arr(i, stdColVBA.stdProcLines)
        End Select
        arrFilter(i, 2) = i
    Next i

    Dim clsSort     As clsSort2DArray
    Set clsSort = New clsSort2DArray
    With clsSort
        Call .SortQuick(arrFilter, 1, True, sdtString)
        arrFilter = .ListArray
    End With

    Dim jCount      As Long
    Dim j           As Long
    jCount = UBound(arr, 2)
    ReDim arrSort(1 To iCount, 1 To jCount)
    For i = 1 To iCount
        For j = 1 To jCount
            arrSort(i, j) = arr(arrFilter(i, 2), j)
        Next j
    Next i
    getArrayCodeBase = arrSort
End Function

' ======================================================
' Private methods: Parsing
' ======================================================
Private Sub ParserControlsUserForm(ByRef oCodeModule As VBIDE.codeModule)
    If oCodeModule.Parent.Type <> vbext_ct_MSForm Then Exit Sub
    Dim oUForm      As UserForm
    Dim sKey        As String
    Set oUForm = oCodeModule.Parent.Designer
    Dim cnt         As control
    Dim sTupeModule As String
    Dim sNameModule As String
    sTupeModule = moduleTypeName(oCodeModule.Parent.Type)
    sNameModule = oCodeModule.Parent.Name
    For Each cnt In oUForm.Controls
        ReDim arrData(1 To 1, 1 To COLUMN_COUNT) As String
        With cnt
            arrData(1, stdColVBA.stdTypeElement) = "UserForms"
            arrData(1, stdColVBA.stdModuleT) = sTupeModule
            arrData(1, stdColVBA.stdModuleName) = sNameModule
            arrData(1, stdColVBA.stdProcType) = "Control"
            arrData(1, stdColVBA.stdProcLines) = TypeName(cnt)
            arrData(1, stdColVBA.stdProcName) = .Name

            On Error Resume Next
            arrData(1, stdColVBA.stdProcVariable) = .Caption
            arrData(1, stdColVBA.stdProcVariable) = .value
            arrData(1, stdColVBA.stdProcDeclaration) = .ControlTipText
            On Error GoTo 0

            sKey = arrData(1, stdColVBA.stdModuleT) & "." & _
                    arrData(1, stdColVBA.stdModuleName) & "." & _
                    arrData(1, stdColVBA.stdProcName)

            If Not mDic.Exists(sKey) Then mDic.Add sKey, arrData
        End With
    Next cnt
End Sub

Private Sub ParserProceduresInModule(ByRef oCodeModule As VBIDE.codeModule, ByRef getCode As Boolean)
    Dim lLineProcF  As Long
    Dim lLineProcE  As Long
    Dim sProcName   As String
    Dim sDeclaration As String
    Dim TypeProc    As vbext_ProcKind
    Dim lCount      As Long
    Dim sKey        As String
    Dim lCounter    As Long
    Dim arrData()   As String

    With oCodeModule
        lLineProcF = .CountOfDeclarationLines + 1
        lCount = .CountOfLines

        Do Until lLineProcF > lCount Or lCounter > MAX_ITERATION_LIMIT
            sProcName = GetProcedureName(oCodeModule, lLineProcF, TypeProc)

            If lLineProcF > 0 And sProcName <> vbNullString Then
                lLineProcE = .ProcCountLines(sProcName, TypeProc)
                sDeclaration = GetProcedureDeclarationString(oCodeModule, sProcName, TypeProc)

                ReDim arrData(1 To 1, 1 To COLUMN_COUNT) As String

                ' Fill the array with procedure data
                arrData(1, stdColVBA.stdTypeElement) = "Procedure"
                arrData(1, stdColVBA.stdModuleT) = moduleTypeName(.Parent.Type)
                arrData(1, stdColVBA.stdModuleName) = .Parent.Name
                arrData(1, stdColVBA.stdProcModifier) = TypeProcedyreModifier(sDeclaration)
                arrData(1, stdColVBA.stdProcType) = TypeProcedyre(sDeclaration)
                arrData(1, stdColVBA.stdProcName) = sProcName
                arrData(1, stdColVBA.stdProcLines) = sProcName
                arrData(1, stdColVBA.stdProcVariable) = GetProcedureVariable(sDeclaration)
                arrData(1, stdColVBA.stdProcDeclaration) = sDeclaration
                arrData(1, stdColVBA.stdProcComment) = GetProcedureComments(oCodeModule, lLineProcF, .CountOfLines)
                If getCode And lLineProcE > 0 Then arrData(1, stdColVBA.stdCode) = clearCodeStrings(.Lines(lLineProcF, lLineProcE))
                ' Generate a unique key for the procedure
                sKey = arrData(1, stdColVBA.stdModuleT) & "." & _
                        arrData(1, stdColVBA.stdModuleName) & "." & _
                        arrData(1, stdColVBA.stdProcModifier) & "." & _
                        arrData(1, stdColVBA.stdProcType) & "." & _
                        sProcName

                If Not mDic.Exists(sKey) Then
                    Call mDic.Add(sKey, arrData)
                    Call ParserparametersInProcedures(arrData, sDeclaration, getCode)
                    Call ParserVariableInProcedures(arrData, .Lines(lLineProcF, lLineProcE), getCode)
                End If
                lLineProcF = lLineProcF + lLineProcE
            Else
                'If procedure not found on current line, move to next
                lLineProcF = lLineProcF + 1
            End If
            lCounter = lCounter + 1
        Loop
    End With
End Sub
Private Sub ParserparametersInProcedures(ByRef arrProc As Variant, ByRef sDeclaration As String, ByRef getCode As Boolean)
    Dim i           As Long
    Dim sParameters As String

    sParameters = clearCodeStrings(sDeclaration)
    If VBA.Right$(sParameters, 2) = "()" Then
        i = VBA.InStrRev(sParameters, " As ") - 1
        If i > 0 Then sParameters = VBA.Left(sParameters, i)
    End If

    i = VBA.InStrRev(sParameters, ")") - 1
    If i < 1 Then Exit Sub
    sParameters = VBA.Left(sParameters, i)

    i = VBA.InStr(1, sParameters, "(")
    If Len(sParameters) - i < 1 Then Exit Sub
    sParameters = VBA.Trim$(VBA.Right$(sParameters, Len(sParameters) - i))

    If VBA.Len(sParameters) = 0 Then Exit Sub
    Dim arrVar      As Variant
    Dim vTokens     As Variant
    Dim lenTokens   As Long
    Dim lIdx        As Long
    Dim sModifier   As String
    Dim sName       As String
    Dim sVarType    As String
    Dim j           As Byte

    arrVar = VBA.Split(sParameters, ", ")
    For i = 0 To UBound(arrVar)
        lIdx = 0: sModifier = vbNullString
        vTokens = VBA.Split(arrVar(i), " ")
        If vTokens(lIdx) = "Optional" Then
            lIdx = lIdx + 1
            sModifier = "Optional "
        End If
        Select Case vTokens(lIdx)
                 Case "ByRef"
                lIdx = lIdx + 1
                sModifier = sModifier & "ByRef"
            Case "ByVal"
                lIdx = lIdx + 1
                sModifier = sModifier & "ByVal"
            Case "ParamArray"
                lIdx = lIdx + 1
                sModifier = sModifier & "ParamArray"
            Case Else
                sModifier = sModifier & "ByRef"
        End Select
        sName = vTokens(lIdx)
        If VBA.Len(sName) > 2 Then
            If VBA.Right$(sName, 2) = "()" Then sName = VBA.Left$(sName, VBA.Len(sName) - 2)
        End If
        lenTokens = UBound(vTokens)
        sVarType = "Variant"
        For j = 0 To lenTokens
            If vTokens(j) = "As" Then
                sVarType = vTokens(j + 1)
                Exit For
            End If
        Next j

        Dim sTypeVar As String
        If sVarType = "Variant" Then
            sTypeVar = typeVariable(sName)
            If sTypeVar <> vbNullString Then sVarType = sTypeVar
        End If

        ReDim arrData(1 To 1, 1 To COLUMN_COUNT) As String
        arrData(1, stdColVBA.stdTypeElement) = "Parametr"
        arrData(1, stdColVBA.stdModuleT) = arrProc(1, stdColVBA.stdModuleT)
        arrData(1, stdColVBA.stdModuleName) = arrProc(1, stdColVBA.stdModuleName)
        arrData(1, stdColVBA.stdProcModifier) = "Private"
        arrData(1, stdColVBA.stdProcType) = arrProc(1, stdColVBA.stdProcType)
        arrData(1, stdColVBA.stdProcName) = arrProc(1, stdColVBA.stdProcName)
        arrData(1, stdColVBA.stdProcLines) = sName
        arrData(1, stdColVBA.stdProcVariable) = sVarType
        arrData(1, stdColVBA.stdProcDeclaration) = sModifier
        If getCode Then
            arrData(1, stdColVBA.stdCode) = VBA.Split(arrProc(1, stdColVBA.stdCode), vbNewLine, 2)(1)
        End If
        Dim sKey    As String
        sKey = arrData(1, stdColVBA.stdModuleT) & "." & _
                arrData(1, stdColVBA.stdModuleName) & "." & _
                arrData(1, stdColVBA.stdProcModifier) & "." & _
                arrData(1, stdColVBA.stdProcType) & "." & _
                arrData(1, stdColVBA.stdProcName) & "." & _
                arrData(1, stdColVBA.stdProcLines)

        If Not mDic.Exists(sKey) Then Call mDic.Add(sKey, arrData)
    Next i
End Sub

Private Sub ParserVariableInProcedures(ByRef arrProc As Variant, ByRef sCode As String, ByRef getCode As Boolean)
    Dim sCodeClear  As String
    sCodeClear = clearCodeStrings(sCode)
    Dim arrLine     As Variant
    Dim vTokens     As Variant
    Dim sName       As String
    Dim sParentName As String
    Dim sVarType    As String
    arrLine = VBA.Split(sCodeClear, vbNewLine)
    Dim i           As Long
    For i = 0 To UBound(arrLine)
        If arrLine(i) <> vbNullString Then
            ReDim arrData(1 To 1, 1 To COLUMN_COUNT) As String
            vTokens = VBA.Split(arrLine(i), " ")
            If vTokens(0) = "Dim" Or vTokens(0) = "Const" Then
                sName = vTokens(1)
                sParentName = VBA.Split(sName, "(")(0)
                If sName Like "*(*" And Not sName Like "*)" Then
                    sName = getName(vTokens, 1)
                End If

                If vTokens(0) = "Const" Then
                    If vTokens(2) = "As" Then
                        sVarType = vTokens(3)
                    Else
                        sVarType = "Variant"
                    End If
                Else
                    If vTokens(UBound(vTokens) - 1) = "As" Then
                        sVarType = vTokens(UBound(vTokens))
                    Else
                        sVarType = "Variant"
                    End If
                End If

                Dim sTypeVar As String
                If sVarType = "Variant" Then
                    sTypeVar = typeVariable(sName)
                    If sTypeVar <> vbNullString Then sVarType = sTypeVar
                End If

                arrData(1, stdColVBA.stdTypeElement) = "Local Variable"
                If vTokens(0) = "Const" Then
                    arrData(1, stdColVBA.stdTypeElement) = "Local Const"
                    arrData(1, stdColVBA.stdProcDeclaration) = VBA.Split(arrLine(i), " = ")(1)
                End If
                arrData(1, stdColVBA.stdModuleT) = arrProc(1, stdColVBA.stdModuleT)
                arrData(1, stdColVBA.stdModuleName) = arrProc(1, stdColVBA.stdModuleName)
                arrData(1, stdColVBA.stdProcModifier) = "Private"
                arrData(1, stdColVBA.stdProcType) = arrProc(1, stdColVBA.stdProcType)
                arrData(1, stdColVBA.stdProcName) = arrProc(1, stdColVBA.stdProcName)
                arrData(1, stdColVBA.stdProcLines) = sParentName
                arrData(1, stdColVBA.stdProcVariable) = sVarType
                If getCode Then
                    arrData(1, stdColVBA.stdProcComment) = arrLine(i)
                    arrData(1, stdColVBA.stdCode) = sCodeClear
                End If

                Dim sKey As String
                sKey = arrData(1, stdColVBA.stdModuleT) & "." & _
                        arrData(1, stdColVBA.stdModuleName) & "." & _
                        arrData(1, stdColVBA.stdProcModifier) & "." & _
                        arrData(1, stdColVBA.stdProcType) & "." & _
                        arrData(1, stdColVBA.stdProcName) & "." & _
                        arrData(1, stdColVBA.stdProcLines)

                If Not mDic.Exists(sKey) Then Call mDic.Add(sKey, arrData)
            End If
        End If
    Next i
End Sub

Private Function GetProcedureDeclarationString(ByRef oCodeModule As VBIDE.codeModule, ByRef sProcName As String, ByVal TypeProc As vbext_ProcKind) As String
    Dim iLineBody   As Long
    Dim sProcBody   As String
    Dim sLineBody   As String

    iLineBody = oCodeModule.ProcBodyLine(sProcName, TypeProc)
    If iLineBody = 0 Then Exit Function

    With oCodeModule
        sLineBody = clearCodeStrings(.Lines(iLineBody, 1))
        ' Handle line continuation
        Do While VBA.Right$(sLineBody, 2) = " _"
            sProcBody = sProcBody & VBA.Left$(sLineBody, VBA.Len(sLineBody) - 2) & " "
            iLineBody = iLineBody + 1
            sLineBody = .Lines(iLineBody, 1)
        Loop
    End With

    sProcBody = WorksheetFunction.Trim(sProcBody & sLineBody)

    ' Trim procedure body (attempt to leave only signature)
    Const KEYWORD_PATTERN As String = "): "
    Dim keywordPos  As Long
    keywordPos = VBA.InStr(1, sProcBody, KEYWORD_PATTERN)

    If keywordPos > 0 Then sProcBody = VBA.Left$(sProcBody, keywordPos)
    GetProcedureDeclarationString = sProcBody
End Function

Private Function GetProcedureVariable(ByVal declarationString As String) As String
    If VBA.Right$(declarationString, 1) = ")" Then
        GetProcedureVariable = vbNullString
    Else
        Const KEYWORD_PATTERN As String = ") As "
        Dim keywordPos As Long
        keywordPos = VBA.InStrRev(declarationString, KEYWORD_PATTERN)
        If keywordPos > 0 Then
            GetProcedureVariable = VBA.mid$(declarationString, keywordPos + VBA.Len(KEYWORD_PATTERN))
        End If
    End If
End Function

Private Function GetProcedureComments(ByRef oCodeModule As VBIDE.codeModule, ByVal lLineProc As Long, ByVal lLineProcEnd As Long) As String
    Dim sLine       As String
    Dim sRes        As String
    Dim i           As Long

    i = lLineProc
    With oCodeModule
        Do Until i > lLineProcEnd
            i = i + 1
            If i > .CountOfLines Then Exit Do

            sLine = VBA.Trim$(.Lines(i, 1))

            If VBA.Left$(sLine, 1) = "'" Then
                If sRes <> vbNullString Then sRes = sRes & vbNewLine
                sRes = sRes & sLine
            Else
                If sLine <> vbNullString Then Exit Do
            End If

            If i > lLineProc + COMMENT_SCAN_LINES Then Exit Do
        Loop
    End With

    GetProcedureComments = sRes
End Function

' ======================================================
' Private methods: Parsing Declaration Section
' ======================================================

Private Sub ParseDeclarationSection(ByRef oCodeModule As VBIDE.codeModule, ByRef getCode As Boolean)
    Dim i           As Long
    Dim sCode       As String
    Dim arrCode     As Variant
    Dim lineIsEnumOrType As Byte
    Dim arrName(1 To 3) As String
    Dim sModuleName As String
    Dim sModuleType As String
    Dim eNumCount As Long

    i = oCodeModule.CountOfDeclarationLines
    If i = 0 Then Exit Sub

    sCode = clearCodeStrings(oCodeModule.Lines(1, i))
    If sCode = vbNullString Then Exit Sub

    arrCode = VBA.Split(sCode, vbNewLine)
    sModuleName = oCodeModule.Parent.Name
    sModuleType = moduleTypeName(oCodeModule.Parent.Type)

    For i = 0 To UBound(arrCode, 1)
        If arrCode(i) <> vbNullString Then
            Call ParseDeclarationLine(sModuleType, sModuleName, arrCode(i), lineIsEnumOrType, arrName, getCode, eNumCount)
        End If
    Next i
End Sub

 Private Sub ParseDeclarationLine( _
        ByRef sModuleType As String, _
        ByRef sModuleName As String, _
        ByVal sLineCode As String, _
        ByRef lineIsEnumOrType As Byte, _
        ByRef arrName() As String, _
        ByRef getCode As Boolean, _
        ByRef eNumCount As Long)

    Dim vTokens     As Variant
    Dim lIdx        As Long
    Dim lLastIdx    As Long
    Dim sModifier   As String
    Dim sType       As String
    Dim sParentName As String
    Dim sName       As String
    Dim sVarType    As String
    Dim sValue      As String
    Dim k           As Long
    Dim m           As Long

    If Len(sLineCode) = 0 Then Exit Sub
    vTokens = VBA.Split(sLineCode, " ")
    lLastIdx = UBound(vTokens, 1)
    lIdx = 0
    Select Case vTokens(lIdx)
             Case "#If", "#Else", "#End": Exit Sub
        Case "End"
            eNumCount = 0
            lineIsEnumOrType = 0
            arrName(1) = vbNullString: arrName(2) = vbNullString: arrName(3) = vbNullString
            Exit Sub
        Case "Option": Exit Sub
        Case "Private": sModifier = "Private": lIdx = lIdx + 1
        Case "Global": sModifier = "Global": lIdx = lIdx + 1
        Case "Public": sModifier = "Public": lIdx = lIdx + 1
    End Select

    If lIdx = 0 Then sModifier = "Private"

    Select Case lineIsEnumOrType
             Case 0    ' Regular module
            Select Case vTokens(lIdx)
                     Case "Enum"
                    lineIsEnumOrType = 1
                    lIdx = lIdx + 1
                    arrName(1) = sModifier: arrName(2) = "Enum": arrName(3) = vTokens(lIdx)
                    sModifier = vbNullString
                    Exit Sub
                Case "Type"
                    lineIsEnumOrType = 2
                    lIdx = lIdx + 1
                    arrName(1) = sModifier: arrName(2) = "Type": arrName(3) = vTokens(lIdx)
                    sModifier = vbNullString
                    Exit Sub
                Case "Const"
                    sType = "Const"
                    lIdx = lIdx + 1
                    sName = vTokens(lIdx)
                    sValue = vTokens(lLastIdx)
                    sVarType = "Variant"
                    For k = lIdx + 1 To lLastIdx
                        If vTokens(k) = "As" Then
                            sVarType = vTokens(k + 1)
                            Exit For
                        End If
                    Next k
                    sParentName = sName
                Case "Declare"
                    For m = 0 To lLastIdx
                        If vTokens(m) = "Lib" Then Exit For
                    Next m
                    If m > 0 Then
                        sType = vTokens(m - 2)
                        sParentName = vTokens(m - 1)
                        sValue = "Lib " & vTokens(m + 1)
                        sName = "API"
                        If sType <> "Sub" And vTokens(lLastIdx - 1) = "As" Then sVarType = vTokens(lLastIdx)
                    End If
                Case "WithEvents"
                    sType = "WithEvents"
                    lIdx = lIdx + 1
                    sParentName = vTokens(lIdx)
                    sName = sParentName
                    sVarType = vTokens(UBound(vTokens))
                Case "Event"
                    sType = "Event"
                    lIdx = lIdx + 1
                    sName = vTokens(lIdx)
                    If sName Like "*(*" And Not sName Like "*)" Then sName = getName(vTokens, lIdx)
                    sParentName = VBA.Split(sName, "(")(0)
                    sName = sParentName
                Case "Dim"
                    sModifier = "Dim"
                    sType = "Variable"
                    lIdx = lIdx + 1
                    sName = vTokens(lIdx)
                    If sName Like "*(*" And Not sName Like "*)" Then sName = getName(vTokens, lIdx)
                    sVarType = "Variant"
                    If lLastIdx >= (lIdx + 2) Then sVarType = vTokens(lIdx + 2)
                    sParentName = VBA.Split(sName, "(")(0)
                Case Else
                    sType = "Variable"
                    sName = vTokens(lIdx)
                    If sName Like "*(*" And Not sName Like "*)" Then sName = getName(vTokens, lIdx)
                    If lLastIdx >= (lIdx + 2) Then
                        If vTokens(lIdx + 1) = "As" Then
                            sVarType = vTokens(lIdx + 2)
                            If sVarType = "New" Then
                                sVarType = vTokens(lIdx + 3)
                            End If
                        End If
                    Else
                        sVarType = "Variant"
                    End If
                    sParentName = VBA.Split(sName, "(")(0)
                    sName = sParentName
            End Select
        Case 1    ' Inside Enum
            sModifier = arrName(1)
            sType = arrName(2)
            sParentName = arrName(3)
            sName = vTokens(0)
            If lLastIdx >= lIdx + 2 Then sValue = vTokens(lIdx + 2)
            If sValue = vbNullString Then
                sValue = eNumCount
                eNumCount = eNumCount + 1
            Else
               If IsNumeric(sValue) Then eNumCount = VBA.CLng(sValue) + 1
            End If
        Case 2    ' Inside Type
            sModifier = arrName(1)
            sType = arrName(2)
            sParentName = arrName(3)
            sName = vTokens(0)
            If sName Like "*(*" And Not sName Like "*)" Then sName = getName(vTokens, lIdx)
            sName = VBA.Split(sName, "(")(0)
            If lLastIdx >= lIdx + 2 Then sVarType = vTokens(lIdx + 2)
            If sVarType = vbNullString Then sVarType = "Variant"
    End Select

    Dim arrResult() As String
    ReDim arrResult(1 To 1, 1 To COLUMN_COUNT) As String

    Dim sTypeVar    As String
    If sVarType = "Variant" Then
        sTypeVar = typeVariable(sName)
        If sTypeVar <> vbNullString Then sVarType = sTypeVar
    End If

    arrResult(1, stdColVBA.stdTypeElement) = "Declaration"
    arrResult(1, stdColVBA.stdModuleT) = sModuleType
    arrResult(1, stdColVBA.stdModuleName) = sModuleName
    arrResult(1, stdColVBA.stdProcModifier) = sModifier
    arrResult(1, stdColVBA.stdProcType) = sType
    arrResult(1, stdColVBA.stdProcName) = sParentName
    arrResult(1, stdColVBA.stdProcLines) = sName
    arrResult(1, stdColVBA.stdProcVariable) = sVarType
    arrResult(1, stdColVBA.stdProcDeclaration) = sValue
    If getCode Then arrResult(1, stdColVBA.stdCode) = sLineCode

    Dim sKey        As String
    sKey = sModuleName & "." & sModifier & "." & sType & "." & sParentName & "." & sName & "." & sVarType & "." & sValue
    If Not mDic.Exists(sKey) Then mDic.Add sKey, arrResult
End Sub

' ======================================================
' Private methods: Utilities and Helper functions
' ======================================================

Private Function getName(ByRef vTokens As Variant, ByRef lIdx As Long)
    Dim sName       As String
    Dim i           As Long
    sName = vTokens(lIdx)
    For i = lIdx + 1 To UBound(vTokens, 1)
        sName = sName & " " & vTokens(i)
        If vTokens(i) Like "*)*" Then
            lIdx = i
            getName = sName
            Exit Function
        End If
    Next i
End Function
' ======================================================
' Initialization and Termination
' ======================================================
Private Sub Class_Terminate()
    Set mDic = Nothing
    Set mVbProj = Nothing
    Set mWB = Nothing
End Sub

Private Sub Class_Initialize()
    Set mDic = New Dictionary
End Sub

Private Sub InitializeContext(ByRef wb As Workbook)
    Set mVbProj = GetVbProject(wb)
End Sub

Private Function GetVbProject(ByRef wb As Workbook) As VBIDE.vbProject
    If mWB Is Nothing Then
        Set mWB = wb
    ElseIf mWB.Name <> wb.Name Then
        Err.Raise vbObjectError + 100, "clsToolsVBACodeStatistics", "Attempted to use object from another workbook. Create a new instance of the class."
    End If

    If mVbProj Is Nothing Then Set mVbProj = mWB.vbProject
    If mVbProj.Protection = vbext_pp_locked Then
        Call MsgBox("Project is password protected!", vbCritical)
    Else
        Set GetVbProject = mVbProj
    End If
End Function

' ======================================================
' Functions: JSON generation (refactoring)
' ======================================================

Public Function getJSONCodeBase(ByRef wb As Workbook) As String
    Call InitializeContext(wb)
    If mVbProj Is Nothing Then
        getJSONCodeBase = "{}"
        Exit Function
    End If

    Call Me.addListModules(wb, False)
    Call Me.addListDeclarations(wb, False)
    Call Me.addListProcs(wb, False)
    Call Me.addListControlsUserForms(wb)

    getJSONCodeBase = "{" & getJSONProject() & getJSONProjectModules() & "}"
End Function

Private Function getJSONProject() As String
    Dim sJSON       As String
    Dim props       As Variant
    Dim i           As Long

    props = Array( _
            Array(JSON_KEY_WORKBOOK_NAME, mWB.Name), _
            Array(JSON_KEY_WORKBOOK_FULL_NAME, mWB.FullName), _
            Array(JSON_KEY_VB_PROJECT_NAME, mWB.vbProject.Name), _
            Array(JSON_KEY_WORKBOOK_AUTHOR, GetOneProp(mWB, "Author")), _
            Array(JSON_KEY_WORKBOOK_LAST_AUTHOR, GetOneProp(mWB, "Last author")), _
            Array(JSON_KEY_WORKBOOK_CREATION_DATE, GetOneProp(mWB, "Creation date")), _
            Array(JSON_KEY_WORKBOOK_LAST_SAVE_DATE, GetOneProp(mWB, "Last save time")), _
            Array(JSON_KEY_COUNT_MODULES, mWB.vbProject.VBComponents.Count) _
            )

    ' Use helper function to build the object
    For i = LBound(props) To UBound(props)
        If i > LBound(props) Then sJSON = sJSON & COMMA

        Dim key     As String: key = CStr(props(i)(0))
        Dim val     As Variant: val = props(i)(1)

        sJSON = sJSON & sQuote(key) & ": "
        If IsNumeric(val) Then
            sJSON = sJSON & CStr(val)
        Else
            sJSON = sJSON & sQuote(EscapeJSON(CStr(val)))
        End If
    Next i

    getJSONProject = sJSON
End Function

Private Function getJSONProjectModules() As String
    Dim i           As Long
    Dim arrVal      As Variant

    ' Buffers for current elements
    Dim sModule     As String
    Dim sDeclaration As String
    Dim sProc       As String
    Dim sParameters As String
    Dim sLocalVariable As String
    Dim sLocalConst As String
    Dim sControls   As String
    Dim sModulesList As String

    arrVal = Me.getArrayCodeBase

    For i = 1 To UBound(arrVal, 1)
        Select Case arrVal(i, stdColVBA.stdTypeElement)
                 Case "Module"
                ' If there is already an open module, close it before starting a new one
                If Len(sModule) > 0 Then
                    sModulesList = AppendModule(sModulesList, sModule, sDeclaration, sProc, sParameters, sLocalVariable, sLocalConst, sControls)
                    ' Reset buffers
                    sModule = vbNullString: sDeclaration = vbNullString: sProc = vbNullString
                    sParameters = vbNullString: sLocalVariable = vbNullString: sLocalConst = vbNullString: sControls = vbNullString
                End If

                ' Start of new module
                sModule = JoinPairs( _
                        sQuote(JSON_KEY_MODULE_NAME) & ": " & sQuote(arrVal(i, stdColVBA.stdModuleName)), _
                        sQuote(JSON_KEY_MODULE_TYPE) & ": " & sQuote(arrVal(i, stdColVBA.stdModuleT)) _
                        )

            Case "Declaration"
                sDeclaration = AddItemJSON(sDeclaration, _
                        "{" & JoinPairs( _
                        sQuote(JSON_KEY_NAME) & ": " & sQuote(IIf(arrVal(i, stdColVBA.stdProcLines) = "API", arrVal(i, stdColVBA.stdProcName), arrVal(i, stdColVBA.stdProcLines))), _
                        sQuote(JSON_KEY_MODIFIER) & ": " & sQuote(arrVal(i, stdColVBA.stdProcModifier)), _
                        sQuote(JSON_KEY_DATA_TYPE) & ": " & sQuote(arrVal(i, stdColVBA.stdProcType)), _
                        IIfQuote(JSON_KEY_VARIABLE_TYPE, arrVal(i, stdColVBA.stdProcVariable)), _
                        IIfQuote(JSON_KEY_VALUE, EscapeJSON(arrVal(i, stdColVBA.stdProcDeclaration))) _
                        ) & "}")

            Case "Procedure"
                ' If there is a previous procedure, finalize it (add parameter and variable arrays)
                If Len(sProc) > 0 Then
                    sProc = FinalizeProcedure(sProc, sParameters, sLocalVariable, sLocalConst)
                    ' Reset local buffers
                    sParameters = vbNullString: sLocalVariable = vbNullString: sLocalConst = vbNullString
                End If
                ' Here we use an approach where sProc accumulates closed procedure objects.
                If Len(sProc) > 0 Then sProc = sProc & COMMA

                sProc = sProc & "{" & JoinPairs( _
                        sQuote(JSON_KEY_NAME) & ": " & sQuote(arrVal(i, stdColVBA.stdProcName)), _
                        sQuote(JSON_KEY_MODIFIER) & ": " & sQuote(arrVal(i, stdColVBA.stdProcModifier)), _
                        sQuote(JSON_KEY_DATA_TYPE) & ": " & sQuote(arrVal(i, stdColVBA.stdProcType)), _
                        IIfQuote(JSON_KEY_VARIABLE_TYPE, EscapeJSON(arrVal(i, stdColVBA.stdProcVariable))) _
                        ) & "}"

            Case "Parametr"
                sParameters = AddItemJSON(sParameters, "{" & JoinPairs( _
                        sQuote(JSON_KEY_NAME) & ": " & sQuote(arrVal(i, stdColVBA.stdProcLines)), _
                        sQuote(JSON_KEY_MODIFIER) & ": " & sQuote(arrVal(i, stdColVBA.stdProcDeclaration)), _
                        sQuote(JSON_KEY_DATA_TYPE) & ": " & sQuote(arrVal(i, stdColVBA.stdProcVariable)) _
                        ) & "}")

            Case "Local Variable"
                sLocalVariable = AddItemJSON(sLocalVariable, "{" & JoinPairs( _
                        sQuote(JSON_KEY_NAME) & ": " & sQuote(arrVal(i, stdColVBA.stdProcLines)), _
                        sQuote(JSON_KEY_DATA_TYPE) & ": " & sQuote(arrVal(i, stdColVBA.stdProcVariable)) _
                        ) & "}")

            Case "Local Const"
                sLocalConst = AddItemJSON(sLocalConst, "{" & JoinPairs( _
                        sQuote(JSON_KEY_NAME) & ": " & sQuote(arrVal(i, stdColVBA.stdProcLines)), _
                        sQuote(JSON_KEY_DATA_TYPE) & ": " & sQuote(arrVal(i, stdColVBA.stdProcVariable)), _
                        sQuote(JSON_KEY_VALUE) & ": " & sQuote(EscapeJSON(arrVal(i, stdColVBA.stdProcDeclaration))) _
                        ) & "}")

            Case "UserForms"
                sControls = AddItemJSON(sControls, "{" & JoinPairs( _
                        sQuote(JSON_KEY_NAME) & ": " & sQuote(arrVal(i, stdColVBA.stdProcName)), _
                        sQuote(JSON_KEY_CONTROL_TYPE) & ": " & sQuote(arrVal(i, stdColVBA.stdProcLines)), _
                        IIfQuote(JSON_KEY_VALUE, EscapeJSON(arrVal(i, stdColVBA.stdProcVariable))), _
                        IIfQuote(JSON_KEY_CONTROL_TIP_TEXT, EscapeJSON(arrVal(i, stdColVBA.stdProcDeclaration))) _
                        ) & "}")
        End Select
    Next i

    If Len(sModule) > 0 Then
        sModulesList = AppendModule(sModulesList, sModule, sDeclaration, sProc, sParameters, sLocalVariable, sLocalConst, sControls)
    End If
    getJSONProjectModules = COMMA & sQuote(JSON_KEY_MODULES) & ": [" & sModulesList & "]"
End Function

Private Function IIfQuote(ByVal sKeyJSON As String, ByVal sVal As String) As String
    IIfQuote = IIf(Len(sVal) > 0, sQuote(sKeyJSON) & ": " & sQuote(sVal), vbNullString)
End Function

Private Function EscapeJSON(ByVal sText As String) As String
    If Len(Trim$(sText)) = 0 Then
        EscapeJSON = vbNullString
        Exit Function
    End If

    Dim sResult     As String
    sResult = sText

    ' Replacement order is important: backslash first, so we don't escape it twice in other replacements
    sResult = Replace(sResult, "\", "\\")
    sResult = Replace(sResult, QUOTE_CHAR, "\" & QUOTE_CHAR)
    sResult = Replace(sResult, vbCr, "\r")
    sResult = Replace(sResult, vbLf, "\n")
    sResult = Replace(sResult, vbTab, "\t")
    sResult = Replace(sResult, vbBack, "\b")
    sResult = Replace(sResult, vbFormFeed, "\f")
    sResult = Replace(sResult, vbNullChar, "\u0000")

    ' Unicode line separators
    sResult = Replace(sResult, ChrW(&H2028), "\u2028")
    sResult = Replace(sResult, ChrW(&H2029), "\u2029")

    EscapeJSON = sResult
End Function

' === HELPER FUNCTIONS ===
' Wraps a string in quotes
Private Function sQuote(ByVal sText As String) As String
    sQuote = QUOTE_CHAR & sText & QUOTE_CHAR
End Function

' Adds an item to a JSON list (array), adding a comma if needed
Private Function AddItemJSON(ByVal sList As String, ByVal sItem As String) As String
    If Len(sList) > 0 Then sList = sList & COMMA
    AddItemJSON = sList & sItem
End Function

' Joins "key: value" pairs with commas
Private Function JoinPairs(ParamArray arrPairs() As Variant) As String
    Dim sRes        As String
    Dim i           As Long
    For i = LBound(arrPairs) To UBound(arrPairs)
        If Len(CStr(arrPairs(i))) > 0 Then
            If Len(sRes) > 0 Then sRes = sRes & COMMA
            sRes = sRes & CStr(arrPairs(i))
        End If
    Next i
    JoinPairs = sRes
End Function

' Finalizes a procedure object: inserts parameter and variable arrays inside the closing brace
Private Function FinalizeProcedure(ByVal sProc As String, ByVal sParams As String, ByVal sLocals As String, ByVal sConsts As String) As String
    ' sProc ends with "}", we need to insert arrays before it
    Dim sContent    As String

    If Len(sParams) > 0 Then sContent = sContent & COMMA & sQuote(JSON_KEY_PARAMETRS) & ": [" & sParams & "]"
    If Len(sLocals) > 0 Then sContent = sContent & COMMA & sQuote(JSON_KEY_LOCAL_VARIABLES) & ": [" & sLocals & "]"
    If Len(sConsts) > 0 Then sContent = sContent & COMMA & sQuote(JSON_KEY_LOCAL_CONSTS) & ": [" & sConsts & "]"

    If Len(sContent) > 0 Then
        ' Remove the last "}" and insert content
        sProc = Left$(sProc, Len(sProc) - 1) & sContent & "}"
    End If

    FinalizeProcedure = sProc
End Function

' Assembles a module into a finished JSON object and adds it to the common list
Private Function AppendModule(ByVal sFullList As String, ByVal sModuleHeader As String, ByVal sDecls As String, ByVal sProcs As String, ByVal sParams As String, ByVal sLocals As String, ByVal sConsts As String, ByVal sControls As String) As String
    Dim sModuleContent As String

    ' Finalize the last procedure in the module, if any
    If Len(sProcs) > 0 Then
        sProcs = FinalizeProcedure(sProcs, sParams, sLocals, sConsts)
        sModuleContent = sQuote(JSON_KEY_PROCEDURES) & ": [" & sProcs & "]"
    End If

    If Len(sDecls) > 0 Then sModuleContent = JoinPairs(sModuleContent, sQuote(JSON_KEY_GLOBAL_DECLARATIONS) & ": [" & sDecls & "]")
    If Len(sControls) > 0 Then sModuleContent = JoinPairs(sModuleContent, sQuote(JSON_KEY_CONTROLS) & ": [" & sControls & "]")

    Dim sFullModule As String
    sFullModule = "{" & JoinPairs(sModuleHeader, sModuleContent) & "}"

    If Len(sFullList) > 0 Then sFullList = sFullList & COMMA
    AppendModule = sFullList & sFullModule
End Function
