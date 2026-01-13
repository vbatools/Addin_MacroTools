Attribute VB_Name = "modToolsDebugOnOff"
Option Explicit
Option Private Module
Option Compare Text
Option Base 1


'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : debugOn - Replaces "'Debug.Print" with "Debug.Print"
'* Created    : 07-06-2023 10:50
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub debugOn()
    Call findeReplaceWordInCodeVBPrj("'Debug.Print", "Debug.Print")
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : debugOff - Replaces "Debug.Print" with "'Debug.Print"
'* Created    : 07-06-2023 10:49
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub debugOff()
    Call findeReplaceWordInCodeVBPrj("Debug.Print", "'Debug.Print")
End Sub

'* * * * * *
'* Sub        : findeReplaceWordInCodeVBPrj - Finds and replaces text in VBA code, depending on selected scope
'* Created    : 07-06-2023 10:49
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                 Description
'*
'* ByVal sFinde As String   : Text to find
'* ByVal sReplace As String : Replacement text
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub findeReplaceWordInCodeVBPrj(ByVal sFinde As String, ByVal sReplace As String)
    Dim vbComp      As VBIDE.VBComponent
    On Error GoTo ErrorHandler
    Select Case WhatIsTextInComboBoxHave(modAddinConst.MENU_TOOLS)
        Case modAddinConst.TYPE_ALL_VBAPROJECT:
            For Each vbComp In Application.VBE.ActiveVBProject.VBComponents
                Call findeReplaceWordInCode(vbComp.CodeModule, sFinde, sReplace)
            Next vbComp
        Case modAddinConst.TYPE_SELECTED_MODULE:
            Call findeReplaceWordInCode(Application.VBE.ActiveCodePane.CodeModule, sFinde, sReplace)
    End Select
    Exit Sub
ErrorHandler:
    Select Case Err.Number
        Case 91:
            Exit Sub
        Case Else:
            Debug.Print "Error in ReBild" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "at line " & Erl
            'Call WriteErrorLog("ReBild")
    End Select
    Err.Clear
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : findeReplaceWordInCode - Finds and replaces text in VBA code
'* Created    : 07-06-2023 10:49
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                             Description
'*
'* ByRef objVBComp As VBIDE.VBComponent : VBA component where the search and replace occurs
'* ByVal sFinde As String               : Text to find
'* ByVal sReplace As String             : Replacement text
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub findeReplaceWordInCode(ByRef objVBComp As VBIDE.CodeModule, ByVal sFinde As String, ByVal sReplace As String)
    Dim sCode       As String
    sCode = GetCodeFromModule(objVBComp)
    If Not sCode Like "*" & sFinde & "*" Then Exit Sub
    sCode = VBA.Replace(sCode, sFinde, sReplace)
    Call SetCodeInModule(objVBComp, sCode)
End Sub
