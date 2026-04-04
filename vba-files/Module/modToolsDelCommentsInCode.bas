Attribute VB_Name = "modToolsDelCommentsInCode"
Option Explicit
Option Private Module

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : delCommentsInCodeVBA - remove comments in code
'* Created    : 19-01-2026 14:20
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub delCommentsInCodeVBA()
    If MsgBox("Delete comments?", vbYesNo + vbQuestion, "Deleting Comments:") = vbNo Then Exit Sub
    Dim moCM        As codeModule
    Dim sCodeVBA    As String
    Dim VBComp      As VBIDE.vbComponent
    On Error GoTo ErrorHandler
    Select Case WhatIsTextInComboBoxHave(modAddinConst.MENU_TOOLS)
        Case modAddinConst.TYPE_ALL_VBAPROJECT:
            For Each VBComp In Application.VBE.ActiveVBProject.VBComponents
                Set moCM = VBComp.codeModule
                sCodeVBA = RemoveCommentsInVBACodeStrings(GetCodeFromModule(moCM))
                If sCodeVBA <> vbNullString Then Call SetCodeInModule(moCM, sCodeVBA)
            Next VBComp
        Case modAddinConst.TYPE_SELECTED_MODULE:
            Set moCM = Application.VBE.ActiveCodePane.codeModule
            Dim iLine As Long
            Call moCM.CodePane.GetSelection(iLine, 0, 0, 0)
            sCodeVBA = RemoveCommentsInVBACodeStrings(GetCodeFromModule(moCM))
            If sCodeVBA <> vbNullString Then
                Call SetCodeInModule(moCM, sCodeVBA)
                Call moCM.CodePane.SetSelection(iLine + 1, 1, iLine + 1, 1)
            End If
    End Select
    Exit Sub
ErrorHandler:
    Select Case Err.Number
        Case 91:
            Exit Sub
        Case Else:
            Call WriteErrorLog("delCommentsInCodeVBA", False)
    End Select
    Err.Clear
End Sub