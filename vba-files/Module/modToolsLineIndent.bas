Attribute VB_Name = "modToolsLineIndent"
Option Explicit
Option Private Module
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* MODULES_REFERENCE.md: L_IndentRoutine - VBA code formatting
'* Created    : 15-09-2019 15:48
'* Created    : 22-03-2023 14:33
'* Author     : VBATools
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

'***************************************************************************
'*
'* PROJECT NAME:    SMART INDENTER
'* AUTHOR:          STEPHEN BULLEN, Office Automation Ltd.
'*
'*                  COPYRIGHT (c) 1999-2004 BY OFFICE AUTOMATION LTD
'*
'* DESCRIPTION:     Adds items to the VBE environment to recreate the indenting
'*                  for the current procedure, module or project.
'*
'* THIS MODULE:     Contains the main procedure to rebuild the code's indenting
'*
'* PROCEDURES:
'*   RebuildModule      Copies the code module to an array for rebuilding and creates a backup
'*   RebuildCodeArray   Main procedure for code formatting
'*   fnFindFirstItem    Check if a code line contains any special service words
'*   CheckLine          Add or remove indent
'*   CopyArrayFromVariant   Convert variant array to string array for faster processing
'*   fnAlignFunction    Find where continuation indent is needed
'*
'***************************************************************************
'*
'* CHANGE HISTORY
'*
'*  DATE        NAME                DESCRIPTION
'*  14/07/1999  Stephen Bullen      Initial version
'*  14/04/2000  Stephen Bullen      Improved algorithm, added options and split out module handling
'*  03/05/2000  Stephen Bullen      Added option to not indent Dims and handle line numbers
'*  24/05/2000  Stephen Bullen      Improved routine for aligning continued lines
'*  27/05/2000  Stephen Bullen      Fix comments with Type/Enum, Rem handling and brackets in strings
'*  04/07/2000  Stephen Bullen      Fix handling of aligned 'As' items and continued lines
'*  24/11/2000  Stephen Bullen      Added maintenance of Members' attributes for VB5 and 6
'*  07/10/2004  Stephen Bullen      Changed to Office Automation
'*  09/10/2004  Stephen Bullen      Bug fixes and more options
'*
'***************************************************************************

'UDT to store Undo information
Public Type UndoRecord
    ModuleObject    As codeModule
    moduleName      As String
    startLine       As Long
    endLine         As Long
    originalLines() As String
    FormattedLines() As String
End Type
Public arrUndo()    As UndoRecord
Const TAB_CHAR      As Integer = 9
Private undoCount   As Integer

' --- Keyword arrays for analysis ---
Private keywordsProcStart() As String
Private keywordsProcEnd() As String
Private keywordsIndentStart() As String
Private keywordsIndentEnd() As String
Private keywordsDeclaration() As String
Private tokensToFind() As String
Private keywordsFunctionAlign() As String

' --- Configuration variables (loaded from settings table) ---
Private configIndentSpaces As Integer
Private configIndentProcedure As Boolean
Private configIndentFirst As Boolean
Private configIndentDim As Boolean
Private configIndentComment As Boolean
Private configIndentCase As Boolean
Private configAlignContinuation As Boolean
Private configAlignIgnoreOperators As Boolean
Private configDebugCol1 As Boolean
Private configAlignDim As Boolean
Private configAlignDimCol As Integer
Private configCompilerCol1 As Boolean
Private configIndentCompiler As Boolean
Private configCommentAlignMode As String
Private configCommentAlignCol As Integer

' --- Service state variables ---
Dim isInitialized   As Boolean
Dim isLineContinued As Boolean
Dim isInsideIfBlock As Boolean
Dim isNoIndentBlock As Boolean
Dim isFirstProcLine As Boolean

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : CutTab - remove all Tabs from VBA code
'* Created    : 08-10-2020 14:08
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub CutTab()
    Dim VBComp      As VBIDE.vbComponent
    On Error GoTo ErrorHandler

    Dim iSelectedRow As Long
    Dim iSelectedCol As Long
    Call Application.VBE.ActiveCodePane.GetSelection(iSelectedRow, iSelectedCol, 1, 1)
    Select Case WhatIsTextInComboBoxHave(modAddinConst.MENU_TOOLS)
        Case modAddinConst.TYPE_ALL_VBAPROJECT:
            For Each VBComp In Application.VBE.ActiveVBProject.VBComponents
                Call TrimLinesTabAndSpase(VBComp.codeModule)
                Next VBComp
        Case modAddinConst.TYPE_SELECTED_MODULE:
            Call TrimLinesTabAndSpase(Application.VBE.ActiveCodePane.codeModule)
            End Select
    If iSelectedRow = 0 Then iSelectedRow = iSelectedRow + 1
    Call Application.VBE.ActiveCodePane.SetSelection(iSelectedRow, iSelectedCol, iSelectedRow, iSelectedCol)
    Exit Sub
ErrorHandler:
    If Err.Number <> 91 Then
        Call WriteErrorLog("CutTab", False)
        End If
    Err.Clear
End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : ReBild - run the VBA code formatting tool
'* Created    : 23-03-2023 10:37
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub ReBild()
    Dim moCM        As codeModule
    Dim cmb_txt     As String
    Dim VBComp      As VBIDE.vbComponent
    On Error GoTo ErrorHandler
    cmb_txt = WhatIsTextInComboBoxHave(modAddinConst.MENU_TOOLS)
    Select Case cmb_txt
        Case TYPE_ALL_VBAPROJECT:
            For Each VBComp In Application.VBE.ActiveVBProject.VBComponents
                Set moCM = VBComp.codeModule
                Call RebuildModule(moCM, moCM.Parent.Name, 1, moCM.CountOfLines, 0)
                Next VBComp
        Case TYPE_SELECTED_MODULE:
            Set moCM = Application.VBE.ActiveCodePane.codeModule
            Call RebuildModule(moCM, moCM.Parent.Name, 1, moCM.CountOfLines, 0)
            End Select
    Exit Sub
ErrorHandler:
    Select Case Err.Number
        Case 91:
            Exit Sub
        Case Else:
            Call WriteErrorLog("ReBild", False)
            End Select
    Err.Clear
End Sub
''''''''''''''''''''''''''''''''''
' Function:   RebuildModule
'
' Comments:   This procedure goes through the lines in a module,
'             rebuilding the code's indenting.
'
' Arguments:  modCode    - The code module to indent
'             sName      - The display name of the item being indented
'             iStartLine - Value giving the line to start indenting from
'             iEndLine   - Value giving the line to end indenting at
'             iProgDone  - Value giving how much indenting has been done in total
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : RebuildModule - format VBA code in a module
'* Created    : 23-03-2023 10:37
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                                     Description
'*
'* ByRef modCode As CodeModule                   : VBA module
'* ByRef sName As String                         : module name
'* ByRef iStartLine As Long                      : starting code line from which formatting begins
'* ByRef iEndline As Long                        : ending code line at which formatting begins
'* ByRef iProgDone As Long                       : number of indents
'* Optional ByRef mbEnableUndo As Boolean = True : ability to undo changes
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub RebuildModule( _
        ByRef modCode As codeModule, _
        ByRef moduleName As String, _
        ByRef startLine As Long, _
        ByRef endLine As Long, _
        Optional ByRef mbEnableUndo As Boolean = True)

    Dim codeLines() As String
    Dim originalLines() As String
    Dim i           As Long

    If endLine = 0 Then Exit Sub
    ReDim codeLines(0 To endLine - startLine)
    ReDim originalLines(0 To endLine - startLine)

       ' Save state for Undo
    If mbEnableUndo Then Call SaveUndoState(modCode, moduleName, startLine, endLine)


       ' Read code
    For i = 0 To endLine - startLine
        codeLines(i) = modCode.Lines(startLine + i, 1)
        originalLines(i) = codeLines(i)
        Next
       ' Main formatting logic
    Call RebuildCodeArray(codeLines)

    For i = 0 To endLine - startLine
        If originalLines(i) <> codeLines(i) Then
            On Error Resume Next
            modCode.ReplaceLine startLine + i, codeLines(i)
            On Error GoTo 0
            End If
        If mbEnableUndo Then arrUndo(undoCount).FormattedLines(i) = codeLines(i)
        Next
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : RebuildCodeArray - modify code in an array
'* Created    : 23-03-2023 10:42
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):         Description
'*
'* ByRef asCodeLines( : array of code lines
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub RebuildCodeArray(ByRef asCodeLines() As String)
       'Variables used for the indenting code
    Dim i As Integer, iGap As Integer, iLineAdjust As Integer
    Dim lLineCount As Long, iCommentStart As Long, iStart As Long, iScan As Long, iDebugAdjust As Integer
    Dim iIndents As Integer, iIndentNext As Integer, iIn As Integer, iOut As Integer
    Dim iFunctionStart As Long, iParamStart As Long
    Dim bInCmt As Boolean, bProcStart As Boolean, bAlign As Boolean, bFirstCont As Boolean
    Dim bAlreadyPadded As Boolean, bFirstDim As Boolean
    Dim sLine As String, sLeft As String, sRight As String, sMatch As String, sItem As String
    Dim iCodeLineNum As Long, sCodeLineNum As String, sOrigLine As String

    Call LoadSettings
    Call InitializeKeywords

       'Flag if the lines are at the top of a procedure
    bProcStart = False
    bFirstDim = False
    bFirstCont = True
       'Loop through all the lines to indent
    For lLineCount = LBound(asCodeLines) To UBound(asCodeLines)
        iLineAdjust = 0
        bAlreadyPadded = False
        iCodeLineNum = -1
        sOrigLine = asCodeLines(lLineCount)
           'Read the line of code to indent
        sLine = Trim$(asCodeLines(lLineCount))
           'If we're not in a continued line, initialise some variables
        If Not (isLineContinued Or bInCmt) Then
            isFirstProcLine = False
            iIndentNext = 0
            iCommentStart = 0
            iIndents = iIndents + iDebugAdjust
            iDebugAdjust = 0
            iFunctionStart = 0
            iParamStart = 0
            i = InStr(1, sLine, " ")
            If i > 0 Then
                If IsNumeric(Left$(sLine, i - 1)) Then
                    iCodeLineNum = val(Left$(sLine, i - 1))
                    sLine = Trim$(mid$(sLine, i + 1))
                    sOrigLine = Space(i) & mid$(sOrigLine, i + 1)
                    End If
                End If
            End If
           'Is there anything on the line?
        If Len(sLine) > 0 Then
               ' Remove leading Tabs
            Do Until Left$(sLine, 1) <> Chr$(TAB_CHAR)
                sLine = mid$(sLine, 2)
                Loop
               ' Add an extra space on the end
            sLine = sLine & " "
            If bInCmt Then
                   'Within a multi-line comment - indent to line up the comment text
                sLine = Space$(iCommentStart) & sLine
                   'Remember if we're in a continued comment line
                bInCmt = Right$(Trim$(sLine), 2) = " _"
                GoTo PTR_REPLACE_LINE
                End If
               'Remember the position of the line segment
            iStart = 1
            iScan = 0
            If isLineContinued And configAlignContinuation Then
                If configAlignIgnoreOperators And Left$(sLine, 2) = ", " Then iParamStart = iFunctionStart - 2
                If configAlignIgnoreOperators And (mid$(sLine, 2, 1) = " " Or Left$(sLine, 2) = ":=") And Left$(sLine, 2) <> ", " Then
                    sLine = Space$(iParamStart - 3) & sLine
                    iLineAdjust = iLineAdjust + iParamStart - 3
                    iScan = iScan + iParamStart - 3
                    Else
                    sLine = Space$(iParamStart - 1) & sLine
                    iLineAdjust = iLineAdjust + iParamStart - 1
                    iScan = iScan + iParamStart - 1
                    End If
                bAlreadyPadded = True
                End If
               'Scan through the line, character by character, checking for
               'strings, multi-statement lines and comments
            Do
                iScan = iScan + 1
                sItem = fnFindFirstItem(sLine, iScan)
                Select Case sItem
                    Case vbNullString
                        iScan = iScan + 1
                           'Nothing found => Skip the rest of the line
                        GoTo PTR_NEXT_PART
                        Case """"
                           'Start of a string => Jump to the end of it
                        iScan = InStr(iScan + 1, sLine, """")
                        If iScan = 0 Then iScan = Len(sLine) + 1
                        Case ": "
                           'A multi-statement line separator => Tidy up and continue
                        If Right$(Left$(sLine, iScan), 6) <> " Then:" Then
                            sLine = Left$(sLine, iScan + 1) & Trim$(mid$(sLine, iScan + 2))
                               'And check the indenting for the line segment
                            CheckLine mid$(sLine, iStart, iScan - 1), iIn, iOut, bProcStart
                            If bProcStart Then bFirstDim = True
                            If iStart = 1 Then
                                iIndents = iIndents - iOut
                                If iIndents < 0 Then iIndents = 0
                                iIndentNext = iIndentNext + iIn
                                Else
                                iIndentNext = iIndentNext + iIn - iOut
                                End If
                            End If
                           'Update the pointer and continue
                        iStart = iScan + 2
                        Case " As "
                           'An " As " in a declaration => Line up to required column
                        If configAlignDim Then
                            bAlign = isNoIndentBlock    'Don't need to check within Type
                            If Not bAlign Then
                                   ' Check if we start with a declaration item
                                For i = LBound(keywordsDeclaration) To UBound(keywordsDeclaration)
                                    sMatch = keywordsDeclaration(i) & " "
                                    If Left$(sLine, Len(sMatch)) = sMatch Then
                                        bAlign = True
                                        Exit For
                                        End If
                                    Next
                                End If
                            If bAlign Then
                                i = InStr(iScan + 3, sLine, " As ")
                                If i = 0 Then
                                       'OK to indent
                                    If configIndentProcedure And bFirstDim And Not configIndentDim And Not isNoIndentBlock Then
                                        iGap = configAlignDimCol - Len(RTrim$(Left$(sLine, iScan)))
                                           'Adjust for a line number at the start of the line
                                        If iCodeLineNum > -1 Then iGap = iGap - Len(CStr(iCodeLineNum)) - 1
                                        Else
                                        iGap = configAlignDimCol - Len(RTrim$(Left$(sLine, iScan))) - iIndents * configIndentSpaces
                                           'Adjust for a line number at the start of the line
                                        If iCodeLineNum > -1 Then
                                            If Len(CStr(iCodeLineNum)) >= iIndents * configIndentSpaces Then
                                                iGap = iGap - (Len(CStr(iCodeLineNum)) - iIndents * configIndentSpaces) - 1
                                                End If
                                            End If
                                        End If
                                    If iGap < 1 Then iGap = 1
                                    Else
                                       'Multiple declarations on the line, so don't space out
                                    iGap = 1
                                    End If
                                   'Work out the new spacing
                                sLeft = RTrim$(Left$(sLine, iScan))
                                sLine = sLeft & Space$(iGap) & mid$(sLine, iScan + 1)
                                   'Update the counters
                                iLineAdjust = iLineAdjust + iGap + Len(sLeft) - iScan
                                iScan = Len(sLeft) + iGap + 3
                                End If
                            Else
                               'Not aligning Dims, so remove any existing spacing
                            iScan = Len(RTrim$(Left$(sLine, iScan)))
                            sLine = RTrim$(Left$(sLine, iScan)) & " " & Trim$(mid$(sLine, iScan + 1))
                            iScan = iScan + 3
                            End If
                        Case "'", "Rem "
                           'The start of a comment => Handle end-of-line comments properly
                        If iScan = 1 Then
                               'New comment at start of line
                            If bProcStart And Not configIndentFirst And Not isNoIndentBlock Then
                                   'No indenting
                                ElseIf configIndentComment Or bProcStart Or isNoIndentBlock Then
                                   'Inside the procedure, so indent to align with code
                                sLine = Space$(iIndents * configIndentSpaces) & sLine
                                iCommentStart = iScan + iIndents * configIndentSpaces
                                ElseIf iIndents > 0 And configIndentProcedure And Not bProcStart Then
                                   'At the top of the procedure, so indent once if required
                                sLine = Space$(configIndentSpaces) & sLine
                                iCommentStart = iScan + configIndentSpaces
                                End If
                            Else
                               'New comment at the end of a line
                               'Make sure it's a proper 'Rem'
                            If sItem = "Rem " And mid$(sLine, iScan - 1, 1) <> " " And mid$(sLine, iScan - 1, 1) <> ":" Then GoTo PTR_NEXT_PART
                               'Check the indenting of the previous code segment
                            CheckLine mid$(sLine, iStart, iScan - 1), iIn, iOut, bProcStart
                            If bProcStart Then bFirstDim = True
                            If iStart = 1 Then
                                iIndents = iIndents - iOut
                                If iIndents < 0 Then iIndents = 0
                                iIndentNext = iIndentNext + iIn
                                Else
                                iIndentNext = iIndentNext + iIn - iOut
                                End If
                               'Get the text before the comment, and the comment text
                            sLeft = Trim$(Left$(sLine, iScan - 1))
                            sRight = Trim$(mid$(sLine, iScan))
                               'Indent the code part of the line
                            If bAlreadyPadded Then
                                sLine = RTrim$(Left$(sLine, iScan - 1))
                                Else
                                If isLineContinued Then
                                    sLine = Space$((iIndents + 2) * configIndentSpaces) & sLeft
                                    Else
                                    If configIndentProcedure And bFirstDim And Not configIndentDim Then
                                        sLine = sLeft
                                        Else
                                        sLine = Space$(iIndents * configIndentSpaces) & sLeft
                                        End If
                                    End If
                                End If
                            isLineContinued = (Right$(Trim$(sLine), 2) = " _")
                               'How do we handle end-of-line comments?
                            Select Case configCommentAlignMode
                                Case "Absolute"
                                    iScan = iScan - iLineAdjust + Len(sOrigLine) - Len(LTrim$(sOrigLine))
                                    iGap = iScan - Len(sLine) - 1
                                    Case "SameGap"
                                    iScan = iScan - iLineAdjust + Len(sOrigLine) - Len(LTrim$(sOrigLine))
                                    iGap = iScan - Len(RTrim$(Left$(sOrigLine, iScan - 1))) - 1
                                    Case "StandardGap"
                                    iGap = configIndentSpaces * 2
                                    Case "AlignInCol"
                                    iGap = configCommentAlignCol - Len(sLine) - 1
                                    End Select
                               'Adjust for a line number at the start of the line
                            If iCodeLineNum > -1 Then
                                Select Case configCommentAlignMode
                                    Case "Absolute", "AlignInCol"
                                        If Len(CStr(iCodeLineNum)) >= iIndents * configIndentSpaces Then
                                            iGap = iGap - (Len(CStr(iCodeLineNum)) - iIndents * configIndentSpaces) - 1
                                            End If
                                        End Select
                                End If
                            If iGap < 2 Then iGap = configIndentSpaces
                            iCommentStart = Len(sLine) + iGap
                               'Put the comment in the required column
                            sLine = sLine & Space$(iGap) & sRight
                            End If
                           'Work out where the text of the comment starts, to align the next line
                        If mid$(sLine, iCommentStart, 4) = "Rem " Then iCommentStart = iCommentStart + 3
                        If mid$(sLine, iCommentStart, 1) = "'" Then iCommentStart = iCommentStart + 1
                        Do Until mid$(sLine, iCommentStart, 1) <> " "
                            iCommentStart = iCommentStart + 1
                            Loop
                        iCommentStart = iCommentStart - 1
                           'Adjust for a line number at the start of the line
                        If iCodeLineNum > -1 Then
                            If Len(CStr(iCodeLineNum)) >= iIndents * configIndentSpaces Then
                                iCommentStart = iCommentStart + (Len(CStr(iCodeLineNum)) - iIndents * configIndentSpaces) + 1
                                End If
                            End If
                           'Remember if we're in a continued comment line
                        bInCmt = Right$(Trim$(sLine), 2) = " _"
                           'Rest of line is comment, so no need to check any more
                        GoTo PTR_REPLACE_LINE
                        Case "Stop ", "Debug.Print ", "Debug.Assert "
                           'A debugging statement - do we want to force to column 1?
                        If configDebugCol1 And iStart = 1 And iScan = 1 Then
                            iLineAdjust = iLineAdjust - (Len(sOrigLine) - LTrim$(Len(sOrigLine)))
                            iDebugAdjust = iIndents
                            iIndents = 0
                            End If
                        Case "#If ", "#ElseIf ", "#Else ", "#End If ", "#Const "
                           'Do we want to force compiler directives to column 1?
                        If configCompilerCol1 And iStart = 1 And iScan = 1 Then
                            iLineAdjust = iLineAdjust - (Len(sOrigLine) - LTrim$(Len(sOrigLine)))
                            iDebugAdjust = iIndents
                            iIndents = 0
                            End If
                        End Select
PTR_NEXT_PART:
                Loop Until iScan > Len(sLine)    'Part of the line
               'Do we have some code left to check?
               '(i.e. a line without a comment or the last segment of a multi-statement line)
            If iStart < Len(sLine) Then
                If Not isLineContinued Then bProcStart = False
                   'Check the indenting of the remaining code segment
                CheckLine mid$(sLine, iStart), iIn, iOut, bProcStart
                If bProcStart Then bFirstDim = True
                If iStart = 1 Then
                    iIndents = iIndents - iOut
                    If iIndents < 0 Then iIndents = 0
                    iIndentNext = iIndentNext + iIn
                    Else
                    iIndentNext = iIndentNext + iIn - iOut
                    End If
                End If
               'Start from the left at each procedure start
            If isFirstProcLine Then iIndents = 0
               ' What about line continuations?  Here, I indent the continued line by
               ' two indents, and check for the end of the continuations.  Note
               ' that Excel won't allow comments in the middle of line continuations
               ' and that comments are treated differently above.
            If isLineContinued Then
                If Not configAlignContinuation Then
                    sLine = Space$((iIndents + 2) * configIndentSpaces) & sLine
                    End If
                Else
                   ' Check if we start with a declaration item
                bAlign = False
                If configIndentProcedure And bFirstDim And Not configIndentDim And Not bProcStart Then
                    For i = LBound(keywordsDeclaration) To UBound(keywordsDeclaration)
                        sMatch = keywordsDeclaration(i) & " "
                        If Left$(sLine, Len(sMatch)) = sMatch Then
                            bAlign = True
                            Exit For
                            End If
                        Next
                    End If
                   'Not a declaration item to left-align, so pad it out
                If Not bAlign Then
                    If Not bProcStart Then bFirstDim = False
                    sLine = Space$(iIndents * configIndentSpaces) & sLine
                    End If
                End If
            isLineContinued = (Right$(Trim$(sLine), 2) = " _")
            End If    'Anything there?
PTR_REPLACE_LINE:
           'Add the code line number back in
        If iCodeLineNum > -1 Then
            sCodeLineNum = CStr(iCodeLineNum)
            If Len(Trim$(Left$(sLine, Len(sCodeLineNum) + 1))) = 0 Then
                sLine = sCodeLineNum & mid$(sLine, Len(sCodeLineNum) + 1)
                Else
                sLine = sCodeLineNum & " " & Trim$(sLine)
                End If
            End If
        asCodeLines(lLineCount) = RTrim$(sLine)
           'If it's not a continued line, update the indenting for the following lines
        If Not isLineContinued Then
            iIndents = iIndents + iIndentNext
            iIndentNext = 0
            If iIndents < 0 Then iIndents = 0
            Else
               'A continued line, so if we're not in a comment and we want smart continuing,
               'work out which to continue from
            If configAlignContinuation And Not bInCmt Then
                If Left$(Trim$(sLine), 2) = "& " Or Left$(Trim$(sLine), 2) = "+ " Then sLine = "  " & sLine
                iFunctionStart = fnAlignFunction(sLine, bFirstCont, iParamStart)
                If iFunctionStart = 0 Then
                    iFunctionStart = (iIndents + 2) * configIndentSpaces
                    iParamStart = iFunctionStart
                    End If
                End If
            End If
        bFirstCont = Not isLineContinued
        Next
End Sub
'
'  Find the first occurrence of one of our key items in the list
'
'    Returns the text of the item found
'    Updates the iFrom parameter to point to the location of the found item
'
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : fnFindFirstItem - find the first occurrence of a searched code string
'* Created    : 23-03-2023 10:43
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):             Description
'*
'* ByRef sLine As String :
'* ByRef iFrom As Long   :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function fnFindFirstItem(ByRef sLine As String, ByRef iFrom As Long) As String
    Dim sItem As String, iFirst As Long, iFound As Long, iItem As Integer
    On Error Resume Next
       'Assume we don't find anything
    iFirst = Len(sLine)
       'Loop through the items to find within the line
    For iItem = LBound(tokensToFind) To UBound(tokensToFind)
           'What to find?
        sItem = tokensToFind(iItem)
           'Is it there?
        iFound = InStr(iFrom, sLine, sItem)
           'Is it before any other items?
        If iFound > 0 And iFound < iFirst Then
            iFirst = iFound
            fnFindFirstItem = sItem
            End If
        Next
       'Update the location of the found item
    iFrom = iFirst
End Function
'  Check the line (segment) to see if it needs in- or out-denting
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : CheckLine - check a code line for moving forward or backward
'* Created    : 23-03-2023 10:44
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                     Description
'*
'* ByVal sLine As String         :
'* ByRef iIndentNext As Integer  :
'* ByRef iOutdentThis As Integer :
'* ByRef bProcStart As Boolean   :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function CheckLine( _
        ByVal sLine As String, _
        ByRef iIndentNext As Integer, _
        ByRef iOutdentThis As Integer, _
        ByRef bProcStart As Boolean)
    Dim i As Integer, j As Integer, sMatch As String
    On Error Resume Next
       'Assume we don't indent or outdent the code
    iIndentNext = 0
    iOutdentThis = 0
       'Tidy up the line
    sLine = Trim$(sLine) & " "
       'We don't check within Type and Enums
    If Not isNoIndentBlock Then
           ' Check for indenting within the code
        For i = LBound(keywordsIndentStart) To UBound(keywordsIndentStart)
            sMatch = keywordsIndentStart(i)
            If (Left$(sLine, Len(sMatch)) = sMatch) And ((mid$(sLine, Len(sMatch) + 1, 1) = " ") Or (mid$(sLine, Len(sMatch) + 1, 1) = ":")) Then
                iIndentNext = iIndentNext + 1
                End If
            Next
           ' Check for out-denting within the code
        For i = LBound(keywordsIndentEnd) To UBound(keywordsIndentEnd)
            sMatch = keywordsIndentEnd(i)
               'Check at start of line for 'real' outdenting
            If (Left$(sLine, Len(sMatch)) = sMatch) And ((mid$(sLine, Len(sMatch) + 1, 1) = " ") Or (mid$(sLine, Len(sMatch) + 1, 1) = ":" And mid$(sLine, Len(sMatch) + 2, 1) <> "=")) Then
                iOutdentThis = iOutdentThis + 1
                End If
            Next
        End If
       'Check procedure-level indenting
    For i = LBound(keywordsProcStart) To UBound(keywordsProcStart)
        sMatch = keywordsProcStart(i)
        If (Left$(sLine, Len(sMatch)) = sMatch) And ((mid$(sLine, Len(sMatch) + 1, 1) = " ") Or (mid$(sLine, Len(sMatch) + 1, 1) = ":" And mid$(sLine, Len(sMatch) + 2, 1) <> "=")) Then
            bProcStart = True
            isFirstProcLine = True
               'Don't indent within Type or Enum constructs
            If Right$(sMatch, 4) = "Type" Or Right$(sMatch, 4) = "Enum" Then
                iIndentNext = iIndentNext + 1
                isNoIndentBlock = True
                ElseIf configIndentProcedure And Not isNoIndentBlock Then
                iIndentNext = iIndentNext + 1
                End If
            Exit For
            End If
        Next
       'Check procedure-level outdenting
    For i = LBound(keywordsProcEnd) To UBound(keywordsProcEnd)
        sMatch = keywordsProcEnd(i)
        If (Left$(sLine, Len(sMatch)) = sMatch) And ((mid$(sLine, Len(sMatch) + 1, 1) = " ") Or (mid$(sLine, Len(sMatch) + 1, 1) = ":" And mid$(sLine, Len(sMatch) + 2, 1) <> "=")) Then
               'Don't indent within Type or Enum constructs
            If Right$(sMatch, 4) = "Type" Or Right$(sMatch, 4) = "Enum" Or configIndentProcedure Then
                iOutdentThis = iOutdentThis + 1
                isNoIndentBlock = False
                End If
            Exit For
            End If
        Next
       'If we're not indenting, no need to consider the special cases
    If isNoIndentBlock Then Exit Function
       ' Treat If as a special case.  If anything other than a comment follows
       ' the Then, we don't indent
    If Left$(sLine, 3) = "If " Or Left$(sLine, 4) = "#If " Or isInsideIfBlock Then
        If isInsideIfBlock Then iIndentNext = 1
           'Strip any strings from the line
        i = InStr(1, sLine, """")
        Do Until i = 0
            j = InStr(i + 1, sLine, """")
            If j = 0 Then j = Len(sLine)
            sLine = Left$(sLine, i - 1) & mid$(sLine, j + 1)
            i = InStr(1, sLine, """")
            Loop
           'And strip comments
        i = InStr(1, sLine, "'")
        If i > 0 Then sLine = Left$(sLine, i - 1)
           ' Do we have a Then statement in the line.  Adding a space on the
           ' end of the test means we can test for Then being both within or
           ' at the end of the line
        sLine = " " & sLine & " "
        i = InStr(1, sLine, " Then ")
           ' Allow for line continuations within the If statement
        isInsideIfBlock = (Right$(Trim$(sLine), 2) = " _")
        If i > 0 Then
               ' If there's something after the Then, we don't indent the If
            If Trim$(mid$(sLine, i + 5)) <> vbNullString Then iIndentNext = 0
               ' No need to check next time around
            isInsideIfBlock = False
            End If
        If isInsideIfBlock Then iIndentNext = 0
        End If
End Function
' Locate the start of the first parameter on the line
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : fnAlignFunction - find the start of the first parameter in a line
'* Created    : 23-03-2023 10:46
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                 Description
'*
'* ByVal sLine As String       :
'* ByRef bFirstLine As Boolean :
'* ByRef iParamStart As Long   :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function fnAlignFunction(ByVal sLine As String, ByRef bFirstLine As Boolean, ByRef iParamStart As Long) As Long
    Dim iLPad As Integer, iCheck As Long, iBrackets As Long, iChar As Long, sMatch As String, iSpace As Integer
    Dim vAlign As Variant, bFound As Boolean, iAlign As Integer
    Dim iFirstThisLine As Integer
    Static coBrackets As Collection
    On Error Resume Next
    ReDim vAlign(1 To 2)
    If bFirstLine Then Set coBrackets = New Collection
       'Convert and numbers at the start of the line to spaces
    iChar = InStr(1, sLine, " ")
    If iChar > 1 Then
        If IsNumeric(Left$(sLine, iChar - 1)) Then
            sLine = mid$(sLine, iChar + 1)
            iLPad = iChar
            End If
        End If
    iLPad = iLPad + Len(sLine) - Len(LTrim$(sLine))
    iFirstThisLine = coBrackets.Count
    sLine = Trim$(sLine)
    iCheck = 1
       'Skip over stuff that we don't want to locate the start off
    For iChar = LBound(keywordsFunctionAlign) To UBound(keywordsFunctionAlign)
        sMatch = keywordsFunctionAlign(iChar)
        If Left$(sLine, Len(sMatch)) = sMatch Then
            iCheck = iCheck + Len(sMatch) + 1
            Exit For
            End If
        Next
    iBrackets = 0
    iSpace = 999
    For iChar = iCheck To Len(sLine)
        Select Case mid$(sLine, iChar, 1)
            Case """"
                   'A String => jump to the end of it
                iChar = InStr(iChar + 1, sLine, """")
                Case "("
                   'Start of another function => remember this position
                vAlign(1) = "("
                vAlign(2) = iChar + iLPad
                coBrackets.Add vAlign
                vAlign(1) = ","
                vAlign(2) = iChar + iLPad + 1
                coBrackets.Add vAlign
                Case ")"
                   'Function finished => Remove back to the previous open bracket
                vAlign = coBrackets(coBrackets.Count)
                Do Until vAlign(1) = "(" Or coBrackets.Count = iFirstThisLine
                    coBrackets.Remove coBrackets.Count
                    vAlign = coBrackets(coBrackets.Count)
                    Loop
                If coBrackets.Count > iFirstThisLine Then coBrackets.Remove coBrackets.Count
                Case " "
                If mid$(sLine, iChar, 3) = " = " Then
                       'Space before an = sign => remember it to align to later
                    bFound = False
                    For iAlign = 1 To coBrackets.Count
                        vAlign = coBrackets(iAlign)
                        If vAlign(1) = "=" Or vAlign(1) = " " Then
                            bFound = True
                            Exit For
                            End If
                        Next
                    If Not bFound Then
                        vAlign(1) = "="
                        vAlign(2) = iChar + iLPad + 2
                        coBrackets.Add vAlign
                        End If
                    ElseIf coBrackets.Count = 0 And iChar < Len(sLine) - 2 Then
                       'Space after a name before the end of the line => remember it for later
                    vAlign(1) = " "
                    vAlign(2) = iChar + iLPad
                    coBrackets.Add vAlign
                    ElseIf iChar > 5 Then
                       'Clear the collection if we find a Then in an If...Then and set the
                       'indenting to align with the bit after the "If "
                    If mid$(sLine, iChar - 5, 6) = " Then " Then
                        Do Until coBrackets.Count <= 1
                            coBrackets.Remove coBrackets.Count
                            Loop
                        End If
                    End If
                Case ","
                   'Start of a new parameter => remember it to align to
                vAlign(1) = ","
                vAlign(2) = iChar + iLPad + 2
                coBrackets.Add vAlign
                Case ":"
                If mid$(sLine, iChar, 2) = ":=" Then
                       'A named paremeter => remember to align to after the name
                    vAlign(1) = ","
                    vAlign(2) = iChar + iLPad + 2
                    coBrackets.Add vAlign
                    ElseIf mid$(sLine, iChar, 2) = ": " Then
                       'A new line section, so clear the brackets
                    Set coBrackets = New Collection
                    iChar = iChar + 1
                    End If
                End Select
        Next
       'If we end with a comma or a named parameter, get rid of all other comma alignments
    If Right$(Trim$(sLine), 3) = ", _" Or Right$(Trim$(sLine), 4) = ":= _" Then
        For iAlign = coBrackets.Count To 1 Step -1
            vAlign = coBrackets(iAlign)
            If vAlign(1) = "," Then
                coBrackets.Remove iAlign
                Else
                Exit For
                End If
            Next
        End If
       'If we end with a "( _", remove it and the space alignment after it
    If Right$(Trim$(sLine), 3) = "( _" Then
        coBrackets.Remove coBrackets.Count
        coBrackets.Remove coBrackets.Count
        End If
    iParamStart = 0
       'Get the position of the unmatched bracket and align to that
    For iAlign = 1 To coBrackets.Count
        vAlign = coBrackets(iAlign)
        If vAlign(1) = "," Then
            iParamStart = vAlign(2)
            ElseIf vAlign(1) = "(" Then
            iParamStart = vAlign(2) + 1
            Else
            iCheck = vAlign(2)
            End If
        Next
    If iCheck = 1 Or iCheck >= Len(sLine) + iLPad - 1 Then
        If coBrackets.Count = 0 And bFirstLine Then
            iCheck = configIndentSpaces * 2 + iLPad
            Else
            iCheck = iLPad
            End If
        End If
    If iParamStart = 0 Then iParamStart = iCheck + 1
    fnAlignFunction = iCheck + 1
End Function

' ==============================================================================
' HELPER FUNCTIONS
' ==============================================================================

' Load settings from Excel table
Private Sub LoadSettings()
    Dim optionsTable As ListObject
    On Error Resume Next
    Set optionsTable = shSettings.ListObjects(TB_OPTIONS_IDEDENT)

    If Err.Number <> 0 Then Exit Sub
    On Error GoTo 0

    With optionsTable.ListColumns(2)
        configIndentSpaces = .Range(2, 1)
        configIndentProcedure = .Range(3, 1)
        configIndentFirst = .Range(4, 1)
        configIndentDim = .Range(5, 1)
        configIndentComment = .Range(6, 1)
        configIndentCase = .Range(7, 1)
        configAlignContinuation = .Range(8, 1)
        configAlignIgnoreOperators = .Range(9, 1)
        configDebugCol1 = .Range(10, 1)
        configAlignDim = .Range(11, 1)
        configAlignDimCol = .Range(12, 1)
        configCompilerCol1 = .Range(13, 1)
        configIndentCompiler = .Range(14, 1)
        configCommentAlignMode = .Range(15, 1)
        configCommentAlignCol = .Range(16, 1)
        End With
End Sub

Private Sub InitializeKeywords()
    Dim vaScope As Variant, vaStatic As Variant, vaType As Variant, vaCombined As Variant
    Dim i As Integer, j As Integer, k As Integer, X As Integer

    If isInitialized Then Exit Sub

       ' Initialize procedure declarations
    vaScope = Array(vbNullString, "Public ", "Private ", "Friend ")
    vaStatic = Array(vbNullString, "Static ")
    vaType = Array("Sub", "Function", "Property Let", "Property Get", "Property Set", "Type", "Enum")

    X = 0
    ReDim vaCombined(0)

    For i = LBound(vaScope) To UBound(vaScope)
        For j = LBound(vaStatic) To UBound(vaStatic)
            For k = LBound(vaType) To UBound(vaType)
                ReDim Preserve vaCombined(X)
                vaCombined(X) = vaScope(i) & vaStatic(j) & vaType(k)
                X = X + 1
                Next k
            Next j
        Next i

    CopyArrayFromVariant keywordsProcStart, vaCombined
    CopyArrayFromVariant keywordsProcEnd, Array("End Sub", "End Function", "End Property", "End Type", "End Enum")

       ' Initialize indent blocks (If, For, etc.)
    If configIndentCompiler Then
        CopyArrayFromVariant keywordsIndentStart, Array("If", "ElseIf", "Else", "#If", "#ElseIf", "#Else", "Select Case", "Case", "With", "For", "Do", "While")
        CopyArrayFromVariant keywordsIndentEnd, Array("ElseIf", "Else", "End If", "#ElseIf", "#Else", "#End If", "Case", "End Select", "End With", "Next", "Loop", "Wend")
        Else
        CopyArrayFromVariant keywordsIndentStart, Array("If", "ElseIf", "Else", "Select Case", "Case", "With", "For", "Do", "While")
        CopyArrayFromVariant keywordsIndentEnd, Array("ElseIf", "Else", "End If", "Case", "End Select", "End With", "Next", "Loop", "Wend")
        End If

       ' Dynamically add Select Case
    If configIndentCase Then
        If keywordsIndentStart(UBound(keywordsIndentStart)) <> "Select Case" Then
            AddToArr keywordsIndentStart, "Select Case"
            AddToArr keywordsIndentEnd, "End Select"
            End If
        End If

    CopyArrayFromVariant keywordsDeclaration, Array("Dim", "Const", "Static", "Public", "Private", "#Const")
    CopyArrayFromVariant tokensToFind, Array("""", ": ", " As ", "'", "Rem ", "Stop ", "#If ", "#ElseIf ", "#Else ", "#End If ", "#Const ", "Debug.Print ", "Debug.Assert ")
    CopyArrayFromVariant keywordsFunctionAlign, Array("Set ", "Let ", "LSet ", "RSet ", "Declare Function", "Declare Sub", "Private Declare Function", "Private Declare Sub", "Public Declare Function", "Public Declare Sub")

    isInitialized = True
End Sub
' ==============================================================================
' UTILITIES AND HELPERS
' ==============================================================================
Private Sub AddToArr(ByRef arr() As String, ByVal value As String)
    ReDim Preserve arr(UBound(arr) + 1)
    arr(UBound(arr)) = value
End Sub

' Convert a Variant array to a string array for faster comparisons
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : CopyArrayFromVariant - converts an array to a string array for faster comparison
'* Created    : 23-03-2023 10:45
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):     Description
'*
'* ByRef asString( :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub CopyArrayFromVariant(ByRef targetArr() As String, ByRef sourceVar As Variant)
    Dim i           As Long
    ReDim targetArr(LBound(sourceVar) To UBound(sourceVar))
    For i = LBound(sourceVar) To UBound(sourceVar)
        targetArr(i) = sourceVar(i)
        Next i
End Sub

' State saving functions (Undo)
Private Sub SaveUndoState(ByRef modCode As codeModule, ByRef sModuleName As String, ByVal lStartLine As Long, ByVal lEndLine As Long)
    undoCount = undoCount + 1
    If undoCount = 1 Then
        ReDim arrUndo(1 To 1)
        Else
        ReDim Preserve arrUndo(1 To undoCount)
        End If

    With arrUndo(undoCount)
        Set .ModuleObject = modCode
        .moduleName = sModuleName
        .startLine = lStartLine
        .endLine = lEndLine
        ReDim .originalLines(0 To lEndLine - lStartLine)
        ReDim .FormattedLines(0 To lEndLine - lStartLine)
        End With

       ' Save originals in the calling code
    Dim i           As Long
    For i = 0 To lEndLine - lStartLine
        arrUndo(undoCount).originalLines(i) = modCode.Lines(lStartLine + i, 1)
        Next
End Sub