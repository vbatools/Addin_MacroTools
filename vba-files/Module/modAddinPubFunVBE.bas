Attribute VB_Name = "modAddinPubFunVBE"
Option Explicit
Option Private Module

Public Enum LineSplits
    LineSplitRemove = 0
    LineSplitKeep = 1
    LineSplitConvert = 2
End Enum


'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : GetSelectControl - get the selected control on a form in the VBE designer
'* Created    : 22-03-2023 16:01
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                             Description
'*
'* Optional bUserForm As Boolean = False :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function GetSelectControl(Optional bUserForm As Boolean = False) As Object
   On Error GoTo ErrorHandler

   If bUserForm Then
       Dim Form    As UserForm
       Set Form = Application.VBE.SelectedVBComponent.Designer
       If Not Form Is Nothing Then
           Set GetSelectControl = Form
           Exit Function
           End If
       Else
       If Application.VBE.ActiveWindow.Type = vbext_wt_Designer Then
           Dim objActiveModule As vbComponent
           Set objActiveModule = getActiveModule()
           If Not objActiveModule Is Nothing Then
               Dim collCntls As Controls
               Set collCntls = objActiveModule.Designer.Selected
               If TypeName(collCntls.Item(0)) = "Frame" And collCntls.Count = 1 Then
                   Dim cnt As control
                   Dim wnd As Object
                   Dim wndAct As Object
                   Set cnt = collCntls.Item(0)
                   Set wndAct = objActiveModule.VBE.ActiveWindow
                   For Each wnd In objActiveModule.VBE.Windows
                       If wnd.Type = vbext_wt_PropertyWindow Then
                           Dim sName As String
                           wnd.Visible = True
                           wndAct.Visible = True
                           sName = VBA.Replace(wnd.Caption, "Properties - ", vbNullString)
                           If sName = objActiveModule.Name Then
                               Debug.Print ">> Inside Frame - only one control can be selected"
                               Set cnt = Nothing
                               Else
                               If sName <> cnt.Name Then Set cnt = objActiveModule.Designer.Controls(sName)
                               End If
                           Set wnd = Nothing
                           Set wndAct = Nothing
                           Exit For
                           End If
                       Next wnd
                   Set GetSelectControl = cnt
                   Else
                   Set GetSelectControl = collCntls
                   End If
               Exit Function
               End If
           End If
       End If

   Exit Function
ErrorHandler:
   Select Case Err.Number
        Case 9:
           Debug.Print ">> For this tool to work, open View -> Properties Window"
        Case Else:
           Call WriteErrorLog("GetSelectControl", False)
           End Select
   Err.Clear
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : getActiveCodePane - get the active code pane
'* Created    : 14-01-2026 14:39
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function getActiveCodePane() As VBIDE.CodePane
   Set getActiveCodePane = Application.VBE.ActiveCodePane
End Function


'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : getActiveModule - get the active VBA module
'* Created    : 22-03-2023 16:00
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function getActiveModule() As vbComponent
   Set getActiveModule = Application.VBE.SelectedVBComponent
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : GetCodeFromModule - get code from module into a string variable
'* Created    : 20-04-2020 18:20
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                             Description
'*
'* ByRef objVBComp As VBIDE.VBComponent : VBA module
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function GetCodeFromModule(ByRef moCM As VBIDE.codeModule) As String
    With moCM
        If .CountOfLines > 0 Then GetCodeFromModule = .Lines(1, .CountOfLines)
        End With
End Function
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : SetCodeInModule - load code from a string variable into a module
'* Created    : 20-04-2020 18:21
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                             Description
'*
'* ByRef objVBComp As VBIDE.VBComponent : VBA module
'* ByVal SCode As String                : string variable
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub SetCodeInModule(ByRef moCM As VBIDE.codeModule, ByVal sCode As String)
    With moCM
        If .CountOfLines > 0 Then Call .DeleteLines(1, .CountOfLines)
        Call .InsertLines(1, VBA.Trim$(sCode))
        End With
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : SelectedLineColumnProcedure - get line and column numbers of selected lines in the VBA module
'* Created    : 08-10-2020 13:48
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function SelectedLineColumnProcedure(ByRef vbCodePane As CodePane) As String
    Dim lStartLine  As Long
    Dim lStartColumn As Long
    Dim lEndLine    As Long
    Dim lEndColumn  As Long

    On Error GoTo ErrorHandler

    With vbCodePane
        .GetSelection lStartLine, lStartColumn, lEndLine, lEndColumn
        SelectedLineColumnProcedure = lStartLine & "|" & lStartColumn & "|" & lEndLine & "|" & lEndColumn
        End With
    Exit Function
ErrorHandler:
    Select Case Err
        Case 91:
            Debug.Print ">> Error! No module activated for code insertion!" & vbNewLine & Err.Number & vbNewLine & Err.Description
        Case Else:
            Call WriteErrorLog("SelectedLineColumnProcedure", False)
            End Select
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : WhatIsTextInComboBoxHave - get the current value in the ComboBox of the VBE editor toolbar
'* Created    : 22-03-2023 14:34
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function WhatIsTextInComboBoxHave(ByVal sTagCombobox As String) As String
    Dim myCommandBar As CommandBar
    Dim cntrl       As CommandBarControl

    Set myCommandBar = Application.VBE.CommandBars(modAddinConst.MENU_TOOLS)
    For Each cntrl In myCommandBar.Controls
        If cntrl.Tag = sTagCombobox Then
            WhatIsTextInComboBoxHave = cntrl.text
            Exit Function
            End If
        Next cntrl
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : TrimLinesTabAndSpase - remove leading tabs and spaces (all lines are left-aligned):
'* Created    : 22-03-2023 16:23
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                             Description
'*
'* ByRef CurCodeModule As VBIDE.CodeModule :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub TrimLinesTabAndSpase(ByRef CurCodeModule As VBIDE.codeModule)
    Dim sLines      As String

    sLines = GetCodeFromModule(CurCodeModule)
    If sLines = vbNullString Then Exit Sub
    Call SetCodeInModule(CurCodeModule, fnTrimLinesTabAndSpase(sLines))
End Sub

Public Function fnTrimLinesTabAndSpase(ByVal strLine As String) As String
    If strLine = vbNullString Then Exit Function
    Dim j           As Long
    Dim arr         As Variant
    Dim sResult     As String
    arr = VBA.Split(strLine, vbNewLine)
    For j = 0 To UBound(arr, 1)
        If sResult <> vbNullString Then sResult = sResult & vbNewLine
        sResult = sResult & VBA.Trim$(arr(j))
        Next j
    fnTrimLinesTabAndSpase = VBA.Trim$(sResult)
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : RemoveCommentsInVBACodeStrings - remove comments
'* Created    : 22-03-2023 16:23
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function RemoveCommentsInVBACodeStrings(ByVal sCode As String) As String
    If sCode = vbNullString Then Exit Function
    If Not sCode Like "*'*" Then
        RemoveCommentsInVBACodeStrings = sCode
        Exit Function
        End If

    Dim i           As Long
    Dim arrLinesCode As Variant
    Dim sLineCode   As String
    Dim posApostr   As Long
    Dim bMultiComment As Boolean

    arrLinesCode = VBA.Split(sCode, vbNewLine)
    For i = 0 To UBound(arrLinesCode)
        sLineCode = arrLinesCode(i)
        posApostr = 1
tryAgain:
        posApostr = VBA.InStr(posApostr, sLineCode, Chr(39))
           'end of multi-line comment
        If bMultiComment And Not VBA.Right$(sLineCode, 2) = " _" Then
            sLineCode = vbNullString
            bMultiComment = False
            End If
        If posApostr > 0 Then
            If VBA.Left$(VBA.LTrim$(sLineCode), 1) = Chr(39) Then
                   'start of multi-line comment
                If VBA.Right$(sLineCode, 2) = " _" Then bMultiComment = True
                sLineCode = vbNullString
                Else
                If CountChrInString(VBA.Left(sLineCode, posApostr - 1), """") Mod 2 = 1 Then
                    posApostr = posApostr + 1
                    GoTo tryAgain
                    Else
                       'start of multi-line comment
                    If VBA.Right$(sLineCode, 2) = " _" Then bMultiComment = True
                    sLineCode = VBA.Left$(sLineCode, posApostr - 1)
                    End If
                End If
            End If
           'middle of multi-line comment
        If bMultiComment And VBA.Right$(sLineCode, 2) = " _" Then sLineCode = vbNullString
        arrLinesCode(i) = sLineCode
        Next i
    RemoveCommentsInVBACodeStrings = VBA.Join(arrLinesCode, vbNewLine)
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : CountChrInString - count how many times the character char appears in the string str
'* Created    : 22-03-2023 16:23
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function CountChrInString(sSTR As String, char As String) As Long
    Dim iResult     As Long
    Dim sParts()    As String

    sParts = Split(sSTR, char)
    iResult = UBound(sParts, 1): If (iResult = -1) Then iResult = 0
    CountChrInString = iResult
End Function

Public Function clearCodeStrings(ByRef sCode As String) As String
    Dim sRes        As String
    sRes = sCode
    sRes = RemoveBreaksLineInCode(sRes)
    sRes = RemoveCommentsInVBACodeStrings(sRes)
    sRes = deleteTwoEmptyCodeStrings(sRes)
    sRes = FormatCodeToMultilineColon(sRes)
    sRes = FormatCodeToMultilineComma(sRes)
    clearCodeStrings = sRes
End Function

Public Function FormatCodeToMultilineComma(ByVal sCode As String) As String
       ' Check for empty string
    If VBA.Len(sCode) = 0 Then
        FormatCodeToMultilineComma = sCode
        Exit Function
        End If

    If sCode Like "*, *" Then
        Dim arr     As Variant
        arr = VBA.Split(sCode, vbNewLine)
        Dim i       As Long
        For i = 0 To UBound(arr, 1)
            If arr(i) Like "*, *" Then
                arr(i) = FormatSingleLineComma(arr(i))
                End If
            Next i
        FormatCodeToMultilineComma = VBA.Join(arr, vbNewLine)
        Else
        FormatCodeToMultilineComma = sCode
        End If
End Function

Public Function FormatSingleLineComma(ByVal sCode As String) As String
    ' Check for empty string
    Dim sRes        As String
    sRes = VBA.Trim$(sCode)
    If VBA.Len(sRes) = 0 Then
        FormatSingleLineComma = sCode
        Exit Function
    End If
    If sCode Like "* Function *" Or sCode Like "* Sub *" Or _
            sCode Like "* Property *" Or sCode Like "* Event *" Or sCode Like "*Const * = " & QUOTE_CHAR & ", " & QUOTE_CHAR Then
        FormatSingleLineComma = sCode
        Exit Function
    End If

    Dim sModif      As String
    Dim arr         As Variant
    sModif = VBA.Split(sRes, " ")(0)
    Select Case sModif
        Case "Dim", "Public", "Private"
            Dim i   As Long
            Dim k   As Long
            Dim iCount As Long
            arr = VBA.Split(sRes, ", ")
            iCount = UBound(arr, 1)
            ReDim arrRes(0 To iCount)
            For i = 0 To iCount
                If arr(i) Like "*(* To *" Then
                    If i = 0 Then
                        arrRes(k) = arr(i)
                    Else
                        k = k + 1
                        arrRes(k) = sModif & " " & arr(i)
                    End If
                ElseIf arr(i) Like "* To *" Then
                    If VBA.Len(arrRes(k)) > 0 Then arrRes(k) = arrRes(k) & ", "
                    arrRes(k) = arrRes(k) & arr(i)
                Else
                    If VBA.Len(arrRes(k)) > 0 Then k = k + 1
                    arrRes(k) = arr(i)
                    If iCount < k Then k = k - 1
                    If Not arrRes(k) Like sModif & " *" And VBA.Len(arrRes(k)) > 0 Then arrRes(k) = sModif & " " & arrRes(k)
                    k = k + 1
                End If
                If iCount < k Then k = k - 1
                If Not arrRes(k) Like sModif & " *" And VBA.Len(arrRes(k)) > 0 Then arrRes(k) = sModif & " " & arrRes(k)
            Next i
            If VBA.Len(arrRes(k)) = 0 Then k = k - 1
            ReDim Preserve arrRes(0 To k)
            sRes = VBA.Join(arrRes, vbNewLine)
    End Select
    FormatSingleLineComma = sRes
End Function

Public Function FormatCodeToMultilineColon(ByVal sCode As String) As String
    If VBA.Len(sCode) = 0 Then
        FormatCodeToMultilineColon = sCode
        Exit Function
        End If

    If sCode Like "*: *" Then
        Dim arr     As Variant
        arr = VBA.Split(sCode, vbNewLine)
        Dim i       As Long
        For i = 0 To UBound(arr, 1)
            If arr(i) Like "*: *" Then
                arr(i) = FormatSingleLineColon(arr(i))
                End If
            Next i
        FormatCodeToMultilineColon = VBA.Join(arr, vbNewLine)
        Else
        FormatCodeToMultilineColon = sCode
        End If
End Function

Public Function FormatSingleLineColon(ByVal sCode As String) As String
       ' Check for empty string
    If VBA.Len(sCode) = 0 Then
        FormatSingleLineColon = sCode
        Exit Function
        End If

    Dim i           As Long
    Dim currentChar As String
    Dim result      As String
    Dim insideString As Boolean
    Dim nextChar    As String

       ' Iterate through each character of the source string
    For i = 1 To Len(sCode)
        currentChar = mid$(sCode, i, 1)

           ' When a quote is encountered, toggle the 'inside string' state flag
        If currentChar = """" Then
            insideString = Not insideString
            result = result & currentChar

               ' When a colon is encountered
            ElseIf currentChar = ":" Then
               ' If we are inside quotes, just add the colon to the result
            If insideString Then
                result = result & currentChar
                   ' If we are OUTSIDE quotes, this is a command separator
                Else
                   ' Check the next character (usually a space follows a colon)
                If i < Len(sCode) Then
                    nextChar = VBA.mid$(sCode, i + 1, 1)
                    If nextChar = " " Then
                           ' Replace ": " with a line break
                        result = result & vbNewLine
                        i = i + 1    ' Skip the space so we don't add it separately
                        ElseIf nextChar = "=" Then
                        result = result & ":"
                        Else
                           ' If there's no space (rare case), just insert a line break
                        result = result & vbNewLine
                        End If
                    Else
                    result = result & vbNewLine
                    End If
                End If
            Else
               ' Regular character, add as-is
            result = result & currentChar
            End If
        Next i
    FormatSingleLineColon = result
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : RemoveBreaksLineInCode - remove code line continuation "* _"
'* Created    : 19-01-2026 14:45
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function RemoveBreaksLineInCode(ByVal sCode As String) As String
    If sCode = vbNullString Then Exit Function
    RemoveBreaksLineInCode = VBA.Replace(sCode, " _" & vbNewLine, " ")
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : GetProcedureDeclaration - get the declaration line
'* Created    : 22-03-2023 15:35
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                                                         Description
'*
'* ByRef CodeMod As VBIDE.CodeModule                                : VBA module
'* ByRef procName As String                                         : procedure name
'* ByRef ProcKind As VBIDE.vbext_ProcKind                           : procedure type
'* Optional ByRef LineSplitBehavior As LineSplits = LineSplitRemove : remove line breaks
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function GetProcedureDeclaration( _
        ByRef codeMod As VBIDE.codeModule, _
        ByRef procName As String, ByRef procKind As VBIDE.vbext_ProcKind, _
        Optional ByRef LineSplitBehavior As LineSplits = LineSplitRemove) As Variant
       '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       ' GetProcedureDeclaration
       ' This return the procedure declaration of ProcName in CodeMod. The LineSplitBehavior
       ' determines what to do with procedure declaration that span more than one line using
       ' the "_" line continuation character. If LineSplitBehavior is LineSplitRemove, the
       ' entire procedure declaration is converted to a single line of text. If
       ' LineSplitBehavior is LineSplitKeep the "_" characters are retained and the
       ' declaration is split with vbNewLine into multiple lines. If LineSplitBehavior is
       ' LineSplitConvert, the "_" characters are removed and replaced with vbNewLine.
       ' The function returns vbNullString if the procedure could not be found.
       '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim LineNum     As Long
    Dim s           As String
    Dim Declaration As String
    On Error Resume Next
    LineNum = codeMod.ProcBodyLine(procName, procKind)
    If Err.Number <> 0 Then Exit Function

    s = codeMod.Lines(LineNum, 1)
    Do While Right$(s, 1) = "_"
        Select Case LineSplitBehavior
            Case LineSplits.LineSplitConvert
                s = Left$(s, Len(s) - 1) & vbNewLine
                Case LineSplits.LineSplitKeep
                s = s & vbNewLine
                Case LineSplits.LineSplitRemove
                s = Left$(s, Len(s) - 1) & " "
                End Select
        Declaration = Declaration & s
        LineNum = LineNum + 1
        s = codeMod.Lines(LineNum, 1)
        Loop
    Declaration = SingleSpace(Declaration & s)
    GetProcedureDeclaration = Declaration
End Function
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : SingleSpace - replace multiple spaces with a single space
'* Created    : 22-03-2023 15:37
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):             Description
'*
'* ByVal sText As String : string to analyze
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function SingleSpace(ByVal sText As String) As String
    Dim pos         As String
    pos = VBA.InStr(1, sText, Space(2), vbBinaryCompare)
    Do Until pos = 0
        sText = VBA.Replace(sText, Space(2), Space(1))
        pos = VBA.InStr(1, sText, Space(2), vbBinaryCompare)
        Loop
    SingleSpace = sText
End Function


'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : VBAIsTrusted - check access to the VBA object model
'* Created    : 22-03-2023 14:33
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function VBAIsTrusted() As Boolean
    On Error GoTo ErrorHandler
    Dim sTxt        As String
    sTxt = Application.VBE.Version
    VBAIsTrusted = True
    Exit Function
ErrorHandler:
    Select Case Err.Number
        Case 1004:
               'If ThisWorkbook.Name = C_Const.NAME_ADDIN & ".xlam" Then
            Call MsgBox("Warning!" & modAddinConst.NAME_ADDIN & vbCrLf & vbNewLine & _
                    "Disabled: [Trust access to the VBA project object model]" & vbCrLf & _
                    "To enable, go to: File->Options->Trust Center->Macro Settings" & _
                    vbCrLf & vbNewLine & "And restart Excel", vbCritical, "Warning:")
        Case Else:
            Call WriteErrorLog("VBAIsTrusted", False)
            End Select
    Err.Clear
    VBAIsTrusted = False
End Function

Public Function GetProcedureName(ByRef oCodeModule As VBIDE.codeModule, ByRef lLineProc As Long, ByRef TypeProc As vbext_ProcKind) As String
    Dim sRes        As String
    Dim procTypes   As Variant
    Dim i           As Integer

    procTypes = Array(vbext_pk_Proc, vbext_pk_Get, vbext_pk_Let, vbext_pk_Set)

    For i = LBound(procTypes) To UBound(procTypes)
        TypeProc = procTypes(i)
        sRes = TryGetProcedureName(oCodeModule, lLineProc, TypeProc)
        If sRes <> vbNullString Then
            GetProcedureName = sRes
            Exit Function
            End If
        Next i
    GetProcedureName = vbNullString
End Function

Public Function TryGetProcedureName(ByRef oCodeModule As VBIDE.codeModule, ByRef lLineProc As Long, ByRef TypeProc As vbext_ProcKind) As String
    On Error Resume Next
    With oCodeModule
        TryGetProcedureName = .ProcOfLine(lLineProc, TypeProc)
        If TryGetProcedureName <> vbNullString Then lLineProc = .ProcStartLine(TryGetProcedureName, TypeProc)
        End With
    On Error GoTo 0
    If lLineProc = 0 Then TryGetProcedureName = vbNullString
End Function

Public Function typeVariable(ByVal sVar As String) As String
    If Len(sVar) < 2 Then Exit Function
    Select Case VBA.Right$(sVar, 1)
        Case "$"
            typeVariable = "String"
            Case "%"
            typeVariable = "Integer"
            Case "&"
            typeVariable = "Long"
            Case "!"
            typeVariable = "Single"
            Case "#"
            typeVariable = "Double"
            Case "@"
            typeVariable = "Currency"
            End Select
End Function
