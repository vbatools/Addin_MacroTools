Attribute VB_Name = "modToolsDelTwoEmptyStrings"
Option Explicit
Option Private Module

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : delTwoEmptyStrings - remove multiple empty lines in VBA code
'* Created    : 23-03-2023 10:12
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub delTwoEmptyStrings()
    Dim moCM        As codeModule
    Dim VBComp      As VBIDE.vbComponent
    On Error GoTo ErrorHandler
    Select Case WhatIsTextInComboBoxHave(modAddinConst.MENU_TOOLS)
        Case modAddinConst.TYPE_ALL_VBAPROJECT:
            For Each VBComp In Application.VBE.ActiveVBProject.VBComponents
                Set moCM = VBComp.codeModule
                Call delEmptyTwoString(moCM)
                Call ReBild
            Next VBComp
        Case modAddinConst.TYPE_SELECTED_MODULE:
            Dim iLine As Long
            Set moCM = Application.VBE.ActiveCodePane.codeModule
            Call moCM.CodePane.GetSelection(iLine, 0, 0, 0)
            Call delEmptyTwoString(moCM)
            Call ReBild
            Call moCM.CodePane.SetSelection(iLine + 1, 1, iLine + 1, 1)
    End Select
    Exit Sub
ErrorHandler:
    Select Case Err.Number
        Case 91:
            Exit Sub
        Case Else:
            Call WriteErrorLog("delTwoEmptyStrings", False)
    End Select
    Err.Clear
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : delEmptyTwoString - remove two consecutive empty lines in VBA code
'* Created    : 23-03-2023 10:12
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                 Description
'*
'* ByRef moCM As CodeModule : VBA code module
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub delEmptyTwoString(ByRef moCM As codeModule)
    Dim sLines      As String
    With moCM
        sLines = GetCodeFromModule(moCM)
        If sLines = vbNullString Then Exit Sub
        Call SetCodeInModule(moCM, deleteTwoEmptyCodeStrings(sLines))
    End With
End Sub

Public Function deleteTwoEmptyCodeStrings(ByVal sCode) As String
    Dim sLines      As String
    Dim sResult     As String
    sLines = fnTrimLinesTabAndSpase(sCode)
    Dim arr         As Variant
    arr = VBA.Split(sLines, vbNewLine)
    Dim j           As Long
    Dim k           As Long
    sResult = vbNullString
    For j = 0 To UBound(arr, 1)
        If arr(j) = vbNullString Then
            For k = j + 1 To UBound(arr, 1)
                If arr(k) <> vbNullString Then
                    j = k - 1
                    Exit For
                End If
            Next k
            sResult = sResult & vbNewLine
        Else
            If sResult <> vbNullString Then sResult = sResult & vbNewLine
            sResult = sResult & arr(j)
        End If
    Next j
    deleteTwoEmptyCodeStrings = sResult
End Function
