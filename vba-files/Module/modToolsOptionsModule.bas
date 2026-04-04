Attribute VB_Name = "modToolsOptionsModule"
Option Explicit
Option Private Module
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : Y_Options - module for creating Options
'* Created    : 17-09-2020 14:35
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : subOptions - launch form for inserting Option into VBA modules
'* Created    : 23-03-2023 10:10
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub subOptionsForm()
    Dim sOptions    As String
    Dim moCM        As codeModule
    Dim VBComp      As VBIDE.vbComponent
    Dim objForm     As frmOptionsModule
    Dim sActiveVBProject As String

    On Error Resume Next
    sActiveVBProject = Application.VBE.ActiveVBProject.FileName
    On Error GoTo 0

    On Error GoTo ErrorHandler
    Set objForm = New frmOptionsModule
    With objForm
        If sActiveVBProject <> vbNullString Then .lbNameProject.Caption = sGetFileName(sActiveVBProject)
        .Show
        If .chOptionExplicit.value Then
            sOptions = "Option Explicit" & vbNewLine
        End If
        If .chOptionPrivate.value Then
            sOptions = sOptions & "Option Private Module" & vbNewLine
        End If
        If .chOptionCompare.value Then
            sOptions = sOptions & "Option Compare Text" & vbNewLine
        End If
        If .chOptionBase.value Then
            sOptions = sOptions & "Option Base 1" & vbNewLine
        End If
        If .chModuleName.value Then
            sOptions = sOptions & vbNewLine & "Private Const MODULE_NAME As String = " & QUOTE_CHAR & "{MODULE_NAME}" & QUOTE_CHAR & vbNewLine
        End If
        If sOptions = vbNullString Then Exit Sub
        sOptions = VBA.Left$(sOptions, VBA.Len(sOptions) - 2)
        If sOptions = vbNullString Then Exit Sub

        If .obtnModule Then
            Set moCM = Application.VBE.ActiveCodePane.codeModule
            Call addString(moCM, IIf(.chModuleName.value, VBA.Replace(sOptions, "{MODULE_NAME}", moCM.Name), sOptions))
        Else
            For Each VBComp In Application.VBE.ActiveVBProject.VBComponents
                Set moCM = VBComp.codeModule
                Call addString(moCM, IIf(.chModuleName.value, VBA.Replace(sOptions, "{MODULE_NAME}", moCM.Name), sOptions))
            Next VBComp
        End If
    End With
    Set objForm = Nothing
    Exit Sub
ErrorHandler:
    Select Case Err.Number
        Case 91:
            Exit Sub
        Case Else:
            Call WriteErrorLog("subOptionsForm", False)
    End Select
    Err.Clear
End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : insertOptionsExplicitAndPrivateModule - quick creation of only Explicit and Private Module options
'* Created    : 23-06-2022 11:20
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub insertOptionsExplicitAndPrivateModule()
    Dim moCM        As codeModule
    Dim VBComp      As VBIDE.vbComponent
    On Error GoTo ErrorHandler
    Select Case WhatIsTextInComboBoxHave(modAddinConst.MENU_TOOLS)
        Case modAddinConst.TYPE_ALL_VBAPROJECT:
            For Each VBComp In Application.VBE.ActiveVBProject.VBComponents
                Set moCM = VBComp.codeModule
                If Not moCM Is Nothing Then Call addString(moCM, "Option Explicit" & vbNewLine & "Option Private Module")
            Next VBComp
        Case modAddinConst.TYPE_SELECTED_MODULE:
            Set moCM = Application.VBE.ActiveCodePane.codeModule
            If Not moCM Is Nothing Then Call addString(moCM, "Option Explicit" & vbNewLine & "Option Private Module")
    End Select
    Exit Sub
ErrorHandler:
    Select Case Err.Number
        Case 91:
            Exit Sub
        Case Else:
            Call WriteErrorLog("insertOptionsExplicitAndPrivateModule", False)
    End Select
    Err.Clear
End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : addString - insert Option lines into a VBA module
'* Created    : 23-03-2023 10:10
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                 Description
'*
'* ByRef moCM As CodeModule : VBA module
'* ByVal sOptions As String : Option lines set in the module
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub addString(ByRef moCM As codeModule, ByVal sOptions As String)
    Dim i           As Long
    Dim sLines      As String
    With moCM
        i = .CountOfDeclarationLines
        If i > 0 Then
            sLines = .Lines(1, i)
            Call .DeleteLines(1, i)

            sLines = VBA.Replace(sLines, "Option Explicit", vbNullString)
            If moCM.Parent.Type = vbext_ct_StdModule Then sLines = VBA.Replace(sLines, "Option Private Module", vbNullString)
            sLines = VBA.Replace(sLines, "Option Base 1", vbNullString)
            sLines = VBA.Replace(sLines, "Option Base 0", vbNullString)
            sLines = VBA.Replace(sLines, "Option Compare Text", vbNullString)
            sLines = VBA.Replace(sLines, "Option Compare Binary", vbNullString)
            sLines = VBA.Replace(sLines, "Private Const MODULE_NAME As String = " & QUOTE_CHAR & moCM.Name & QUOTE_CHAR, vbNullString)
            If .Parent.Type <> vbext_ct_StdModule Then sOptions = VBA.Replace(sOptions, "Option Private Module", vbNullString)
            sLines = deleteTwoEmptyCodeStrings(sLines)
        End If
        If moCM.Parent.Type <> vbext_ct_StdModule Then sOptions = VBA.Replace(sOptions, vbNewLine & "Option Private Module", vbNullString)
        If sLines <> vbNullString Then sOptions = sOptions & vbNewLine & sLines
        Call .InsertLines(1, sOptions)
    End With
End Sub