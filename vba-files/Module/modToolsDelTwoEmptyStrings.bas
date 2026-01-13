Attribute VB_Name = "modToolsDelTwoEmptyStrings"
Option Explicit
Option Private Module
Option Compare Text
Option Base 1

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : delTwoEmptyStrings - Removes duplicate empty lines in VBA code
'* Created    : 23-03-2023 10:12
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub delTwoEmptyStrings()
    Dim moCM        As CodeModule
    Dim vbComp      As VBIDE.VBComponent
    On Error GoTo ErrorHandler
    Select Case WhatIsTextInComboBoxHave(modAddinConst.MENU_TOOLS)
        Case modAddinConst.TYPE_ALL_VBAPROJECT:
            For Each vbComp In Application.VBE.ActiveVBProject.VBComponents
                Set moCM = vbComp.CodeModule
                Call delEmptyTwoString(moCM)
                Call ReBild
            Next vbComp
        Case modAddinConst.TYPE_SELECTED_MODULE:
            Set moCM = Application.VBE.ActiveCodePane.CodeModule
            Call delEmptyTwoString(moCM)
            Call ReBild
    End Select
    Exit Sub
ErrorHandler:
    Select Case Err.Number
        Case 91:
            Exit Sub
        Case Else:
            Debug.Print "Error in delTwoEmptyStrings" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "at line " & Erl
            'Call WriteErrorLog("delTwoEmptyStrings")
    End Select
    Err.Clear
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : delEmptyTwoString - Removes double empty lines in VBA code
'* Created    : 23-03-2023 10:12
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                 Description
'*
'* ByRef moCM As CodeModule :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub delEmptyTwoString(ByRef moCM As CodeModule)
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

