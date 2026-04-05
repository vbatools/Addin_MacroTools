Attribute VB_Name = "modToolsSnipets"
Option Explicit
Option Private Module

Private Enum tbSnipetsCol
    tbCodeGrup = 1
    tbCodeSnipet
    tbCode
    tbDiscr
    tbClassName
    tbClassCode
    tbFormaName
    tbFormaFRM
    tbFormaFRX
End Enum

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : InsertCodeFromSnippet - insert snippet into VBA code module
'* Created    : 22-03-2023 14:57
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub InsertCodeFromSnippet()
    Dim objActCodePane As VBIDE.CodePane
    Set objActCodePane = getActiveCodePane()

    ' Check: is code pane active
    If objActCodePane Is Nothing Then
        Debug.Print ">> No VBA module activated for code insertion!"
        Exit Sub
    End If

    Dim startLine   As Long
    Dim endLine     As Long
    Dim startCol    As Long
    Dim endCol      As Long
    Dim currentLine As String
    Dim indentationLevel As Long

    ' Get selection and current line
    With objActCodePane
        .GetSelection startLine, startCol, endLine, endCol
        currentLine = .codeModule.Lines(startLine, 1)

        If currentLine = vbNullString Then
            Debug.Print ">> Nothing selected!"
            Exit Sub
        End If

        ' Calculate indentation level (number of spaces on the left)
        indentationLevel = Len(currentLine) - Len(LTrim$(currentLine))
    End With

    currentLine = Trim$(currentLine)

    Dim replacementArg As String
    Dim rawToken    As String
    Dim cleanToken  As String
    Dim lastSpacePos As Long
    Dim lineTokens  As Variant
    Dim tokenCount  As Long

    lastSpacePos = InStrRev(currentLine, " ")

    ' Parse line logic to extract token and replacement argument
    If lastSpacePos > 0 Then
        lineTokens = Split(currentLine, " ")
        tokenCount = UBound(lineTokens)

        Select Case tokenCount
                 Case 0    ' Single word (technically impossible with lastSpacePos > 0, but kept for safety)
                rawToken = currentLine
            Case 1    ' Two words: first is token, second is replacement argument
                replacementArg = lineTokens(1)
                rawToken = lineTokens(0)
            Case Else    ' More than two words: last word is token, rest is prefix
                rawToken = lineTokens(tokenCount)
        End Select

        ' Form the prefix (text before token)
        If lastSpacePos > 0 And tokenCount > 1 Then
            currentLine = Left$(currentLine, lastSpacePos)
        Else
            currentLine = vbNullString
        End If
    Else
        ' No spaces: entire line is the token
        rawToken = currentLine
        currentLine = vbNullString
    End If

    ' Clean token from prefix (e.g., "Form.Show" -> "Show")
    cleanToken = rawToken
    If rawToken Like "*.*" Then
        cleanToken = Right$(rawToken, Len(rawToken) - InStrRev(rawToken, "."))
    End If

    ' Get snippets table
    Dim snippetsTable As Variant
    snippetsTable = getArrayTBSnipets()

    If Not IsArray(snippetsTable) Then Exit Sub

    ' Search for snippet in table by cleaned token
    Dim rowIndex    As Long
    rowIndex = findeValueInTabel(cleanToken, snippetsTable, tbSnipetsCol.tbCodeSnipet)

    If rowIndex = -1 Then Exit Sub

    ' Insert dependencies (modules and forms) and code
    If addSnipetModules(snippetsTable(rowIndex, tbSnipetsCol.tbClassName), snippetsTable(rowIndex, tbSnipetsCol.tbClassCode)) And _
            addSnipetForms(snippetsTable(rowIndex, tbSnipetsCol.tbFormaName), snippetsTable(rowIndex, tbSnipetsCol.tbFormaFRM), snippetsTable(rowIndex, tbSnipetsCol.tbFormaFRX)) Then

        Dim codeToInsert As String
        codeToInsert = snippetsTable(rowIndex, tbSnipetsCol.tbCode)

        ' Build final line: prefix + snippet code
        codeToInsert = currentLine & codeToInsert

        ' Substitute argument in place of @1 placeholder
        If replacementArg <> vbNullString Then codeToInsert = Replace(codeToInsert, "@1", replacementArg)

        ' Apply indentation
        If indentationLevel > 0 Then codeToInsert = AddSpaceCode(codeToInsert, indentationLevel)

        ' Insert code into editor
        With objActCodePane
            .codeModule.ReplaceLine startLine, codeToInsert
            .SetSelection startLine + 1, startCol, startLine + 1, startCol
        End With
    End If
End Sub

Private Function addSnipetForms(ByVal sFormaName As String, ByVal sFormaFRM As String, ByVal sFormaFRX As String) As Boolean
    Dim arrFormaName As Variant
    Dim arrFormaFRM As Variant
    Dim arrFormaFRX As Variant

    arrFormaName = VBA.Split(sFormaName, ";")
    arrFormaFRM = VBA.Split(sFormaFRM, ";")
    arrFormaFRX = VBA.Split(sFormaFRX, ";")

    Dim i           As Integer
    Dim iCount      As Integer
    iCount = UBound(arrFormaName, 1)

    If iCount <> UBound(arrFormaFRM, 1) Or iCount <> UBound(arrFormaFRX, 1) Then
        Debug.Print ">> Error creating form - number of names and code blocks" & vbCrLf & _
                "Inserting Form:" & vbCrLf & vbTab & _
                "[" & iCount + 1 & "][" & sFormaName & "]" & vbCrLf & vbTab & _
                "[" & UBound(arrFormaFRM, 1) + 1 & "][" & sFormaFRM & "]" & vbCrLf & vbTab & _
                "[" & UBound(arrFormaFRX, 1) + 1 & "][" & sFormaFRX & "]"
        Exit Function
    End If

    For i = 0 To iCount
        Call addSnipetForm(arrFormaName(i), arrFormaFRM(i), arrFormaFRX(i))
    Next i
    addSnipetForms = True
End Function

Private Sub addSnipetForm(ByVal sFormaName As String, ByVal sFormaFRM As String, ByVal sFormaFRX As String)
    Dim sPath       As String
    sPath = Environ$("Temp") & Application.PathSeparator & sFormaName

    sFormaFRM = getCodeFromShape(sFormaFRM)
    sFormaFRX = getCodeFromShape(sFormaFRX)

    If sFormaFRM <> vbNullString And sFormaFRX <> vbNullString Then
        Call base64ToFile(sFormaFRM, sPath & ".frm")
        Call base64ToFile(sFormaFRX, sPath & ".frx")
        Call Application.VBE.ActiveVBProject.VBComponents.Import(FileName:=sPath & ".frm")
    End If

    If FileHave(sPath & ".frm", vbNormal) Then Call Kill(sPath & ".frm")
    If FileHave(sPath & ".frx", vbNormal) Then Call Kill(sPath & ".frx")
End Sub

Private Function addSnipetModules(ByVal sClassName As String, ByVal sClassCode As String) As Boolean
    Dim arrName     As Variant
    Dim arrCode     As Variant

    arrName = VBA.Split(sClassName, ";")
    arrCode = VBA.Split(sClassCode, ";")

    Dim i           As Integer
    Dim iCount      As Integer
    iCount = UBound(arrName, 1)

    If iCount <> UBound(arrCode, 1) Then
        Debug.Print ">> Error creating module - number of names and code blocks" & vbCrLf & _
                "Inserting Module:" & vbCrLf & vbTab & _
                "[" & iCount + 1 & "][" & sClassName & "]" & vbCrLf & vbTab & _
                "[" & UBound(arrCode, 1) + 1 & "][" & sClassCode & "]"
        Exit Function
    End If

    For i = 0 To iCount
        Call addSnipetModule(arrName(i), arrCode(i))
    Next i
    addSnipetModules = True
End Function

Private Sub addSnipetModule(ByVal sNameModule As String, ByVal sNameShapeOrCode As String)
    If sNameModule = vbNullString Or sNameShapeOrCode = vbNullString Then Exit Sub
    Dim sCode       As String
    sCode = getCodeFromShape(sNameShapeOrCode)
    If sCode = vbNullString Then sCode = sNameShapeOrCode

    If sNameShapeOrCode Like "TB_CLS_*" Then
        Call AddModuleToProject(Application.VBE.ActiveVBProject, sNameModule, vbext_ct_ClassModule, sCode, False)
    ElseIf sNameShapeOrCode Like "TB_MOD_*" Then
        Call AddModuleToProject(Application.VBE.ActiveVBProject, sNameModule, vbext_ct_StdModule, sCode, False)
    Else
        Call AddModuleToProject(Application.VBE.ActiveVBProject, sNameModule, vbext_ct_StdModule, sCode, False)
    End If
End Sub

Private Function getCodeFromShape(ByRef sNameShape As String) As String
    On Error GoTo endfun
    getCodeFromShape = shSettings.Shapes(sNameShape).TextFrame2.TextRange.text
    Exit Function
endfun:
    Debug.Print ">> Shape with code not found [" & sNameShape & "]"
    Err.Clear
End Function

Private Function findeValueInTabel(ByVal sValue As Variant, ByRef arr As Variant, ByVal iCol As Integer) As Long
    Dim i           As Long
    Dim iCount      As Long
    iCount = UBound(arr, 1)
    For i = 1 To iCount
        If VBA.LCase$(arr(i, iCol)) = VBA.LCase$(sValue) Then
            findeValueInTabel = i
            Exit Function
        End If
    Next i
    findeValueInTabel = -1
    Debug.Print ">> Snippet not found, selected: " & sValue
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : AddSpaceCode - function to add spaces to a code line
'* Created    : 22-03-2023 14:59
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):             Description
'*
'* ByRef code As String : code line
'* ByRef spac As Long   : number of spaces
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function AddSpaceCode(ByRef sCode As String, ByVal countLeftSpase As Long) As String
    Dim strarr()    As String
    Dim newCode As String, spaceStr As String
    Dim i           As Long
    Dim iCount      As Long
    spaceStr = VBA.Space(countLeftSpase)
    strarr = Split(sCode, Chr$(10))
    iCount = UBound(strarr)
    For i = 0 To iCount
        If i = iCount Then
            newCode = newCode & spaceStr & strarr(i)
        Else
            newCode = newCode & spaceStr & strarr(i) & Chr$(10)
        End If
    Next i
    AddSpaceCode = newCode
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : AddSnippetEnumModule - insert snippet enum module
'* Created    : 08-10-2020 14:01
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub AddSnippetEnumModule()
    Call AddModuleToProject(Application.VBE.ActiveVBProject, modAddinConst.MODULE_NAME_SNIPETS, vbext_ct_StdModule, AddEnumCode(), False)
End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : DeleteSnippetEnumModule - delete snippet enum module
'* Created    : 08-10-2020 14:01
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub DeleteSnippetEnumModule()
    Call DeleteModuleToProject(Application.VBE.ActiveVBProject, modAddinConst.MODULE_NAME_SNIPETS)
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : AddEnumCode - generate VBA code for snippet autocomplete
'* Created    : 14-01-2026 14:04
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function AddEnumCode() As String
    Dim arrTB       As Variant
    Dim sCode       As String
    Dim sFun        As String
    Dim sType       As String
    Dim sDiscriptions As String
    Dim countSpace  As Byte
    Dim i           As Long

    arrTB = getArrayTBSnipets()
    If Not IsArray(arrTB) Then Exit Function

    sDiscriptions = "'SNIPPET GROUP CODES | GROUP DESCRIPTION"
    countSpace = VBA.Len(VBA.Split(sDiscriptions, " | ")(0))

    For i = 1 To UBound(arrTB)
        If sType <> arrTB(i, tbSnipetsCol.tbCodeGrup) Then
            sType = arrTB(i, tbSnipetsCol.tbCodeGrup)
            If sCode <> vbNullString Then sCode = sCode & "End Type" & vbCrLf & vbCrLf
            sCode = sCode & "Private Type " & sType & vbCrLf
            sFun = sFun & "Public Function " & sType & "() As " & sType & ": End Function" & vbCrLf
            sDiscriptions = sDiscriptions & vbCrLf & "'" & sType & VBA.Space(countSpace - VBA.Len(sType) - 1) & " | "
        End If
        sCode = sCode & Space(4) & "'" & arrTB(i, tbSnipetsCol.tbDiscr) & vbCrLf & Space(4) & arrTB(i, tbSnipetsCol.tbCodeSnipet) & " as byte" & vbCrLf
    Next i
    AddEnumCode = getCommentVBATools() & vbNewLine & sDiscriptions & vbCrLf & vbCrLf & sCode & "End Type" & vbCrLf & vbCrLf & sFun
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : getArrayTBSnipets - get snippet table array
'* Created    : 14-01-2026 14:27
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function getArrayTBSnipets() As Variant
    On Error GoTo endfun
    getArrayTBSnipets = shSettings.ListObjects(modAddinConst.TB_SNIPETS).DataBodyRange.Value2
    Exit Function
endfun:
    Debug.Print ">> Error in AddModuleToProject" & vbCrLf & Err.Number & vbCrLf & Err.Description & vbCrLf & "at line" & Erl
    Err.Clear
End Function
