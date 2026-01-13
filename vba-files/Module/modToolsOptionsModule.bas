Attribute VB_Name = "modToolsOptionsModule"
Option Explicit
Option Private Module
Option Compare Text
Option Base 1
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : Y_Options - Module for handling Options
'* Created    : 17-09-2020 14:35
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *



'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : subOptions - Shows the Options form in VBA editor
'* Created    : 23-03-2023 10:10
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub subOptionsForm()
    Dim sOptions    As String
    Dim moCM        As CodeModule
    Dim vbComp      As VBIDE.VBComponent
    Dim objForm     As frmOptionsModule
    Dim sActiveVBProject As String

    On Error Resume Next
    sActiveVBProject = Application.VBE.ActiveVBProject.Filename
    On Error GoTo 0

    On Error GoTo ErrorHandler
    Set objForm = New frmOptionsModule
    With objForm
        If sActiveVBProject <> vbNullString Then .lbNameProject.Caption = sGetFileName(sActiveVBProject)
        .Show
        If .chOptionExplicit.Value Then
            sOptions = "Option Explicit" & vbNewLine
        End If
        If .chOptionPrivate.Value Then
            sOptions = sOptions & "Option Private Module" & vbNewLine
        End If
        If .chOptionCompare.Value Then
            sOptions = sOptions & "Option Compare Text" & vbNewLine
        End If
        If .chOptionBase.Value Then
            sOptions = sOptions & "Option Base 1" & vbNewLine
        End If
        If sOptions = vbNullString Then Exit Sub
        sOptions = VBA.Left$(sOptions, VBA.Len(sOptions) - 2)
        If sOptions = vbNullString Then Exit Sub

        If .obtnModule Then
            Set moCM = Application.VBE.ActiveCodePane.CodeModule
            Call addString(moCM, sOptions)
        Else
            For Each vbComp In Application.VBE.ActiveVBProject.VBComponents
                Set moCM = vbComp.CodeModule
                Call addString(moCM, sOptions)
            Next vbComp
        End If
    End With
    Set objForm = Nothing
    Exit Sub
ErrorHandler:
    Select Case Err.Number
        Case 91:
            Exit Sub
        Case Else:
            Debug.Print "Error in addOptions" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "at line " & Erl
            'Call WriteErrorLog("addOptions")
    End Select
    Err.Clear
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : insertOptionsExplicitAndPrivateModule - Inserts Option Explicit and Private Module statements
'* Created    : 23-06-2022 11:20
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub insertOptionsExplicitAndPrivateModule()
    Dim moCM        As CodeModule
    Dim vbComp      As VBIDE.VBComponent
    On Error GoTo ErrorHandler
    Select Case WhatIsTextInComboBoxHave(modAddinConst.MENU_TOOLS)
        Case modAddinConst.TYPE_ALL_VBAPROJECT:
            For Each vbComp In Application.VBE.ActiveVBProject.VBComponents
                Set moCM = vbComp.CodeModule
                If Not moCM Is Nothing Then Call addString(moCM, "Option Explicit" & vbNewLine & "Option Private Module")
            Next vbComp
        Case modAddinConst.TYPE_SELECTED_MODULE:
            Set moCM = Application.VBE.ActiveCodePane.CodeModule
            If Not moCM Is Nothing Then Call addString(moCM, "Option Explicit" & vbNewLine & "Option Private Module")
    End Select
    Exit Sub
ErrorHandler:
    Select Case Err.Number
        Case 91:
            Exit Sub
        Case Else:
            Debug.Print "Error in insertOptionsExplicitAndPrivateModule" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "at line " & Erl
            'Call WriteErrorLog("insertOptionsExplicitAndPrivateModule")
    End Select
    Err.Clear
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : addString - Adds Option statements to VBA modules
'* Created    : 23-03-2023 10:10
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                 Description
'*
'* ByRef moCM As CodeModule : VBA Module
'* ByVal sOptions As String : Option statements to insert
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub addString(ByRef moCM As CodeModule, ByVal sOptions As String)
    Dim i           As Long
    Dim sLines      As String
    With moCM
        i = .CountOfDeclarationLines
        If i > 0 Then
            sLines = .Lines(1, i)
            Call .DeleteLines(1, i)

            sLines = VBA.Replace(sLines, "Option Explicit", vbNullString)
            sLines = VBA.Replace(sLines, "Option Private Module", vbNullString)
            sLines = VBA.Replace(sLines, "Option Base 1", vbNullString)
            sLines = VBA.Replace(sLines, "Option Base 0", vbNullString)
            sLines = VBA.Replace(sLines, "Option Compare Text", vbNullString)
            sLines = VBA.Replace(sLines, "Option Compare Binary", vbNullString)
            If .Parent.Type <> vbext_ct_StdModule Then sOptions = VBA.Replace(sOptions, "Option Private Module", vbNullString)
            sLines = deleteTwoEmptyCodeStrings(sLines)
        End If
        Call .InsertLines(1, sOptions & vbNewLine & sLines)
    End With
End Sub
