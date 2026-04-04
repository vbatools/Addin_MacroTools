Attribute VB_Name = "modToolsDebugOnOff"
Option Explicit
Option Private Module

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : debugOn - replace "'Debug.Print" with "Debug.Print"
'* Created    : 07-06-2023 10:50
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub debugOn()
    Call findeReplaceWordInCodeVBPrj("'Debug.Print", "Debug.Print")
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : debugOff - replace "Debug.Print" with "'Debug.Print"
'* Created    : 07-06-2023 10:49
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub debugOff()
    Call findeReplaceWordInCodeVBPrj("Debug.Print", "'Debug.Print")
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : findeReplaceWordInCodeVBPrj - search and replace text in VBA code, in all modules or selected module
'* Created    : 07-06-2023 10:49
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                 Description
'*
'* ByVal sFinde As String   : search text
'* ByVal sReplace As String : replacement text
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub findeReplaceWordInCodeVBPrj(ByVal sFinde As String, ByVal sReplace As String)
    Dim VBComp      As VBIDE.vbComponent
    On Error GoTo ErrorHandler
    Select Case WhatIsTextInComboBoxHave(modAddinConst.MENU_TOOLS)
        Case modAddinConst.TYPE_ALL_VBAPROJECT:
            For Each VBComp In Application.VBE.ActiveVBProject.VBComponents
                Call findeReplaceWordInCode(VBComp.codeModule, sFinde, sReplace)
            Next VBComp
        Case modAddinConst.TYPE_SELECTED_MODULE:
            Call findeReplaceWordInCode(Application.VBE.ActiveCodePane.codeModule, sFinde, sReplace)
    End Select
    Exit Sub
ErrorHandler:
    Select Case Err.Number
        Case 91:
            Exit Sub
        Case Else:
            Call WriteErrorLog("findeReplaceWordInCodeVBPrj", False)
    End Select
    Err.Clear
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : findeReplaceWordInCode - search and replace text in VBA code
'* Created    : 07-06-2023 10:49
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                             Description
'*
'* ByRef objVBComp As VBIDE.VBComponent : VBA component to search for text to replace
'* ByVal sFinde As String               : search text
'* ByVal sReplace As String             : replacement text
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub findeReplaceWordInCode(ByRef objVBComp As VBIDE.codeModule, ByVal sFinde As String, ByVal sReplace As String)
    Dim sCode       As String
    sCode = GetCodeFromModule(objVBComp)
    If Not sCode Like "*" & sFinde & "*" Then Exit Sub
    sCode = VBA.Replace(sCode, sFinde, sReplace)
    Call SetCodeInModule(objVBComp, sCode)
End Sub