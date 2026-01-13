Attribute VB_Name = "modToolsLineNumbers"
Option Explicit
Option Private Module
Option Compare Text
Option Base 1
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : K_AddNumbersLine - Module for adding line numbers
'* Created    : 22-03-2023 15:40
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Public Enum vbLineNumbers_LabelTypes
    vbLabelColon    ' 0
    vbLabelTab    ' 1
End Enum

Private Enum vbLineNumbers_ScopeToAddLineNumbersTo
    vbScopeAllProc    ' 1
    vbScopeThisProc    ' 2
End Enum

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : AddLineNumbers_ - Adds line numbers to VBA modules
'* Created    : 22-03-2023 15:40
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub AddLineNumbers_()
    On Error GoTo ErrorHandler
    Dim vbComp      As VBIDE.VBComponent
    Select Case WhatIsTextInComboBoxHave(modAddinConst.MENU_TOOLS)
        Case modAddinConst.TYPE_ALL_VBAPROJECT:
            For Each vbComp In Application.VBE.ActiveVBProject.VBComponents
                AddLineNumbers vbCompObj:=vbComp, LabelType:=vbLabelColon, AddLineNumbersToEmptyLines:=True, AddLineNumbersToEndOfProc:=True, Scope:=vbScopeAllProc
            Next vbComp
        Case modAddinConst.TYPE_SELECTED_MODULE:
            AddLineNumbers vbCompObj:=Application.VBE.ActiveCodePane.CodeModule.Parent, LabelType:=vbLabelColon, AddLineNumbersToEmptyLines:=True, AddLineNumbersToEndOfProc:=True, Scope:=vbScopeAllProc
    End Select
    Exit Sub
ErrorHandler:
    Select Case Err.Number
        Case 91:
            Exit Sub
        Case Else:
            Debug.Print "Error in AddLineNumbers_" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "at line " & Erl
            'Call WriteErrorLog("AddLineNumbers_")
    End Select
    Err.Clear
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : RemoveLineNumbersPublic - Removes line numbers
'* Created    : 2-03-2023 15:40
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub RemoveLineNumbersPublic()
    On Error GoTo ErrorHandler
    Dim vbComp      As VBIDE.VBComponent
    Select Case WhatIsTextInComboBoxHave(modAddinConst.MENU_TOOLS)
        Case modAddinConst.TYPE_ALL_VBAPROJECT:
            For Each vbComp In Application.VBE.ActiveVBProject.VBComponents
                RemoveLineNumbers vbCompObj:=vbComp, LabelType:=vbLabelColon
                RemoveLineNumbers vbCompObj:=vbComp, LabelType:=vbLabelTab
            Next vbComp
        Case modAddinConst.TYPE_SELECTED_MODULE:
            RemoveLineNumbers vbCompObj:=Application.VBE.ActiveCodePane.CodeModule.Parent, LabelType:=vbLabelColon
            RemoveLineNumbers vbCompObj:=Application.VBE.ActiveCodePane.CodeModule.Parent, LabelType:=vbLabelTab
    End Select
    Exit Sub
ErrorHandler:
    Select Case Err.Number
        Case 91:
            Exit Sub
        Case Else:
            Debug.Print "Error in RemoveLineNumbersPublic" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "at line " & Erl
            'Call WriteErrorLog("RemoveLineNumbersPublic")
    End Select
    Err.Clear
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : AddLineNumbers - Adds line numbers to VBA modules
'* Created    : 22-03-2023 15:41
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                                             Description
'*
'* ByVal vbCompObj As VBIDE.VBComponent                 : VBA Component
'* ByVal LabelType As vbLineNumbers_LabelTypes          : Type of label for line numbers
'* ByVal AddLineNumbersToEmptyLines As Boolean          : Add to blank line numbers
'* ByVal AddLineNumbersToEndOfProc As Boolean           : Add at the end of procedures
'* ByVal Scope As vbLineNumbers_ScopeToAddLineNumbersTo :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub AddLineNumbers( _
        ByVal vbCompObj As VBIDE.VBComponent, _
        ByVal LabelType As vbLineNumbers_LabelTypes, _
        ByVal AddLineNumbersToEmptyLines As Boolean, _
        ByVal AddLineNumbersToEndOfProc As Boolean, _
        ByVal Scope As vbLineNumbers_ScopeToAddLineNumbersTo)
    ' USAGE RULES
    ' DO NOT MIX LABEL TYPES FOR LINE NUMBERS! IF ADDING LINE NUMBERS AS COLON TYPE, ANY LINE NUMBERS AS VBTAB TYPE MUST BE REMOVE BEFORE, AND RECIPROCALLY ADDING LINE NUMBERS AS VBTAB TYPE
    Dim i           As Long
    Dim procName    As String
    Dim startOfProcedure As Long
    Dim lengthOfProcedure As Long
    Dim endOfProcedure As Long
    Dim bodyOfProcedure As Long
    Dim countOfProcedure As Long
    Dim prelinesOfProcedure As Long
    Dim PreviousIndentAdded As Long
    Dim strLine     As String
    Dim temp_strLine As String
    Dim new_strLine As String
    Dim tupe_procedure As vbext_ProcKind
    Dim InProcBodyLines As Boolean
    Dim FlagSelect  As Boolean
    With vbCompObj.CodeModule

        If Scope = vbScopeAllProc Then
            For i = 1 To .CountOfLines - 1
                strLine = .Lines(i, 1)
                If FlagSelect Then
                    FlagSelect = False
                    GoTo NextLine
                End If
                If strLine Like "*Select Case *" Then FlagSelect = True
                procName = .ProcOfLine(i, tupe_procedure)    ' Type d'argument ByRef incompatible ~~> Requires VBIDE library as a Reference for the VBA Project
                If procName <> vbNullString Then
                    startOfProcedure = .ProcStartLine(procName, tupe_procedure)
                    bodyOfProcedure = .ProcBodyLine(procName, tupe_procedure)
                    countOfProcedure = .ProcCountLines(procName, tupe_procedure)
                    prelinesOfProcedure = bodyOfProcedure - startOfProcedure
                    'postlineOfProcedure = ??? not directly available since endOfProcedure is itself not directly available.
                    lengthOfProcedure = countOfProcedure - prelinesOfProcedure    ' includes postlinesOfProcedure !
                    'endOfProcedure = ??? not directly available, each line of the proc must be tested until the End statement is reached. See below.
                    If endOfProcedure <> 0 And startOfProcedure < endOfProcedure And i > endOfProcedure Then
                        GoTo NextLine
                    End If
                    If i = bodyOfProcedure Then InProcBodyLines = True
                    If bodyOfProcedure < i And i < startOfProcedure + countOfProcedure Then
                        If Not (.Lines(i - 1, 1) Like "* _") Then
                            InProcBodyLines = False
                            PreviousIndentAdded = 0
                            If Trim$(strLine) = vbNullString And Not AddLineNumbersToEmptyLines Then GoTo NextLine
                            If IsProcEndLine(vbCompObj, i) Then
                                endOfProcedure = i
                                If AddLineNumbersToEndOfProc Then
                                    Call IndentProcBodyLinesAsProcEndLine(vbCompObj, LabelType, endOfProcedure, tupe_procedure)
                                Else
                                    GoTo NextLine
                                End If
                            End If
                            If LabelType = vbLabelColon Then
                                If HasLabel(strLine, vbLabelColon) Then strLine = RemoveOneLineNumber(.Lines(i, 1), vbLabelColon)
                                If Not HasLabel(strLine, vbLabelColon) Then
                                    temp_strLine = strLine
                                    On Error Resume Next
                                    .ReplaceLine i, CStr(i) & ":" & strLine
                                    On Error GoTo 0
                                    new_strLine = .Lines(i, 1)
                                    If Len(new_strLine) = Len(CStr(i) & ":" & temp_strLine) Then
                                        PreviousIndentAdded = Len(CStr(i) & ":")
                                    Else
                                        PreviousIndentAdded = Len(CStr(i) & ": ")
                                    End If
                                End If
                            ElseIf LabelType = vbLabelTab Then
                                If Not HasLabel(strLine, vbLabelTab) Then strLine = RemoveOneLineNumber(.Lines(i, 1), vbLabelTab)
                                If Not HasLabel(strLine, vbLabelColon) Then
                                    temp_strLine = strLine
                                    On Error Resume Next
                                    .ReplaceLine i, CStr(i) & vbTab & strLine
                                    On Error GoTo 0
                                    PreviousIndentAdded = Len(strLine) - Len(temp_strLine)
                                End If
                            End If
                        Else
                            If Not InProcBodyLines Then
                                If LabelType = vbLabelColon Then
                                    On Error Resume Next
                                    .ReplaceLine i, Space(PreviousIndentAdded) & strLine
                                    On Error GoTo 0
                                ElseIf LabelType = vbLabelTab Then
                                    On Error Resume Next
                                    .ReplaceLine i, Space(4) & strLine
                                    On Error GoTo 0
                                End If
                            Else
                            End If
                        End If
                    End If
                End If
NextLine:
            Next i
        ElseIf AddLineNumbersToEmptyLines And Scope = vbScopeThisProc Then
            'TODO selected prosedure
        End If

    End With
End Sub

'* * * *
'* Function   : IsProcEndLine - Checks if the line is the end of a procedure
'* Created    : 22-03-2023 15:47
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                             Description
'*
'* ByVal vbCompObj As VBIDE.VBComponent : VBA Component
'* ByVal lLine As Long                  : Line number
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function IsProcEndLine( _
        ByVal vbCompObj As VBIDE.VBComponent, _
        ByVal lLine As Long) As Boolean
    With vbCompObj.CodeModule
        If Trim$(.Lines(lLine, 1)) Like "End Sub*" _
                Or Trim$(.Lines(lLine, 1)) Like "End Function*" _
                Or Trim$(.Lines(lLine, 1)) Like "End Property*" _
                Then IsProcEndLine = True
    End With
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : IndentProcBodyLinesAsProcEndLine - Indents lines in procedure body according to end line
'* Created    : 22-03-2023 15:55
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                                 Description
'*
'* ByVal vbCompObj As VBIDE.VBComponent        : VBA Component
'* ByVal LabelType As vbLineNumbers_LabelTypes : Type of label for line numbers
'* ByVal ProcEndLine As Long                   : Line number of end of procedure
'* ByVal VBEXT As vbext_ProcKind               : Procedure type
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub IndentProcBodyLinesAsProcEndLine( _
        ByVal vbCompObj As VBIDE.VBComponent, _
        ByVal LabelType As vbLineNumbers_LabelTypes, _
        ByVal ProcEndLine As Long, _
        ByVal VBEXT As vbext_ProcKind)
    Dim procName    As String
    Dim bodyOfProcedure As Long
    Dim j           As Long
    Dim endOfProcedure As Long
    Dim strEnd      As String
    Dim strLine     As String
    With vbCompObj.CodeModule
        procName = .ProcOfLine(ProcEndLine, VBEXT)
        bodyOfProcedure = .ProcBodyLine(procName, VBEXT)
        endOfProcedure = ProcEndLine
        strEnd = .Lines(endOfProcedure, 1)
        j = bodyOfProcedure
        If j = 1 Then j = 2
        Do Until Not .Lines(j - 1, 1) Like "* _" And j <> bodyOfProcedure
            strLine = .Lines(j, 1)
            If LabelType = vbLabelColon Then
                If Mid$(strEnd, Len(CStr(endOfProcedure)) + 1 + 1 + 1, 1) = " " Then
                    On Error Resume Next
                    .ReplaceLine j, Space(Len(CStr(endOfProcedure)) + 1) & strLine
                    On Error GoTo 0
                Else
                    On Error Resume Next
                    .ReplaceLine j, Space(Len(CStr(endOfProcedure)) + 2) & strLine
                    On Error GoTo 0
                End If
            ElseIf LabelType = vbLabelTab Then
                If endOfProcedure < 1000 Then
                    On Error Resume Next
                    .ReplaceLine j, Space(4) & strLine
                    On Error GoTo 0
                Else
                    Debug.Print "Maximum supported line number is 999 lines for this procedure type."
                End If
            End If
            j = j + 1
        Loop
    End With
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : RemoveLineNumbers - Removes line numbers from VBA modules
'* Created    : 22-03-2023 15:45
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                                 Description
'*
'* ByVal vbCompObj As VBIDE.VBComponent        : VBA Component
'* ByVal LabelType As vbLineNumbers_LabelTypes : Type of label for line numbers
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub RemoveLineNumbers(ByVal vbCompObj As VBIDE.VBComponent, ByVal LabelType As vbLineNumbers_LabelTypes)
    Dim i           As Long
    Dim RemovedChars_previous_i As Long
    Dim procName    As String
    Dim InProcBodyLines As Boolean
    Dim tupe_procedure As vbext_ProcKind
    With vbCompObj.CodeModule
        'Debug.Print ("nr of lines = " & .CountOfLines & vbNewLine & "Procname = " & procName)
        'Debug.Print ("nr of lines REMEMBER MUST BE LARGER THAN 7! = " & .CountOfLines)
        For i = 1 To .CountOfLines
            procName = .ProcOfLine(i, tupe_procedure)
            If procName <> vbNullString Then
                If i > 1 Then
                    'Debug.Print ("Line " & i & " is a body line " & .ProcBodyLine(procName, tupe_procedure))
                    If i = .ProcBodyLine(procName, tupe_procedure) Then InProcBodyLines = True
                    If Not .Lines(i - 1, 1) Like "* _" Then
                        'Debug.Print (InProcBodyLines)
                        InProcBodyLines = False
                        'Debug.Print ("recoginized a line that should be substituted: " & i)
                        'Debug.Print ("about to replace " & .Lines(i, 1) & vbNewLine & " with: " & RemoveOneLineNumber(.Lines(i, 1), LabelType) & vbNewLine & " with label type: " & LabelType)
                        On Error Resume Next
                        .ReplaceLine i, RemoveOneLineNumber(.Lines(i, 1), LabelType)
                        On Error GoTo 0
                    Else
                        If InProcBodyLines Then
                            ' do nothing
                            'Debug.Print i
                        Else
                            On Error Resume Next
                            .ReplaceLine i, Mid$(.Lines(i, 1), RemovedChars_previous_i + 1)
                            On Error GoTo 0
                        End If
                    End If
                End If
            Else
            End If
        Next i
    End With
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : RemoveOneLineNumber - Removes one line number from a string
'* Created    : 22-03-2023 15:43
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                                 Description
'*
'* ByVal aString As String                     : String
'* ByVal LabelType As vbLineNumbers_LabelTypes : Label type Tab or Colon
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function RemoveOneLineNumber(ByVal aString As String, ByVal LabelType As vbLineNumbers_LabelTypes) As Variant
    RemoveOneLineNumber = aString
    If LabelType = vbLabelColon Then
        If aString Like "#:*" Or aString Like "##:*" Or aString Like "###:*" Or aString Like "####:*" Then
            RemoveOneLineNumber = Mid$(aString, 1 + InStr(1, aString, ":", vbTextCompare))
            If Left$(RemoveOneLineNumber, 2) Like " [! ]*" Then RemoveOneLineNumber = Mid$(RemoveOneLineNumber, 2)
        End If
    ElseIf LabelType = vbLabelTab Then
        If aString Like "#   *" Or aString Like "##  *" Or aString Like "### *" Or aString Like "#### *" Then RemoveOneLineNumber = Mid$(aString, 5)
        If aString Like "#" Or aString Like "##" Or aString Like "###" Or aString Like "####" Then RemoveOneLineNumber = vbNullString
    End If
    If RemoveOneLineNumber Like "*Function *" Or RemoveOneLineNumber Like "*Sub *" _
            Or RemoveOneLineNumber Like "*Property Set *" Or RemoveOneLineNumber Like "*Property Get *" Or RemoveOneLineNumber Like "*Property Let *" Then
        RemoveOneLineNumber = RemoveLeadingSpaces(RemoveOneLineNumber)
    End If
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : HasLabel - Checks if a string contains a label at the beginning
'* Created    : 22-03-2023 15:43
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                                 Description
'*
'* ByVal aString As String                     :
'* ByVal LabelType As vbLineNumbers_LabelTypes :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function HasLabel(ByVal aString As String, ByVal LabelType As vbLineNumbers_LabelTypes) As Boolean
    If LabelType = vbLabelColon Then HasLabel = InStr(1, aString & ":", ":") < InStr(1, aString & " ", " ")
    If LabelType = vbLabelTab Then
        HasLabel = Mid$(aString, 1, 4) Like "#   " Or Mid$(aString, 1, 4) Like "##  " Or Mid$(aString, 1, 4) Like "### " Or Mid$(aString, 1, 5) Like "#### "
    End If
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : RemoveLeadingSpaces - Removes leading spaces from a string
'* Created    : 22-03-2023 15:42
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):             Description
'*
'* ByVal aString As String :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function RemoveLeadingSpaces(ByVal aString As String) As String
    Do Until Left$(aString, 1) <> " "
        aString = Mid$(aString, 2)
    Loop
    RemoveLeadingSpaces = aString
End Function
