Attribute VB_Name = "modToolsLineIndent"
Option Explicit
Option Private Module
Option Compare Text
Option Base 1
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : L_IndentRoutine - VBA Indentation Module
'* Created    : 15-09-2019 15:48
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'*
'* PROJECT NAME:    SMART INDENTER
'* AUTHOR:          STEPHEN BULLEN, Office Automation Ltd.
'*
'*                  COPYRIGHT © 1999-2004 BY OFFICE AUTOMATION LTD
'*
'* CONTACT:         stephen@oaltd.co.uk
'* WEB SITE:        http://www.oaltd.co.uk
'*
'* DESCRIPTION:     Adds items to the VBE environment to recreate the indenting
'*                  for the current procedure, module or project.
'*
'* THIS MODULE:     Contains the main procedure to rebuild the code's indenting
'*
'* PROCEDURES:
'*   RebuildModule      Rebuilds the indenting for a procedure or module
'*   RebuildCodeArray   Rebuilds the indenting for a code array
'*   fnFindFirstItem    Finds the first occurrence of a key item in a line
'*   CheckLine          Checks a line for indenting requirements
'*   ArrayFromVariant   Converts a variant array to a string array for comparison
'*   fnAlignFunction    Locates the start of the first parameter in a line
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
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
'*  24/1/2000  Stephen Bullen      Added maintenance of Members' attributes for VB5 and 6
'*  07/10/2004  Stephen Bullen      Changed to Office Automation
'*  09/10/2004  Stephen Bullen      Bug fixes and more options
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'UDT to store Undo information
Public Type uUndo
    oMod            As CodeModule
    sName           As String
    lStartLine      As Long
    lEndLine        As Long
    asOriginal()    As String
    asIndented()    As String
End Type
Public pauUndo()    As uUndo
Const miTAB         As Integer = 9
Public piUndoCount  As Integer
'Variable arrays to hold the code items to look for
'Variant arrays to hold the code items to look for
Dim masInProc() As String, masInCode() As String, masOutProc() As String, masOutCode() As String
Dim masDeclares() As String, masLookFor() As String, masFnAlign() As String
'Variables for storing configuration settings
Dim mbIndentProc As Boolean, mbIndentCmt As Boolean, mbIndentCase As Boolean, mbAlignCont As Boolean, mbIndentDim As Boolean
Dim mbIndentFirst As Boolean, mbAlignDim As Boolean, mbDebugCol1 As Boolean, mbEnableUndo As Boolean
Dim miIndentSpaces As Integer, miEOLAlignCol As Integer, miAlignDimCol As Integer, mbCompilerStuffCol1 As Boolean
Dim mbIndentCompilerStuff As Boolean, mbAlignIgnoreOps As Boolean
'Variables to hold operational information
Dim mbInitialised As Boolean, mbContinued As Boolean, mbInIf As Boolean, mbNoIndent As Boolean, mbFirstProcLine As Boolean
Dim msEOLComment    As String

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : CutTab - Removes Tabs from VBA code
'* Created    : 08-10-2020 14:08
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub CutTab()
    Dim vbComp      As VBIDE.VBComponent
    On Error GoTo ErrorHandler
    Select Case WhatIsTextInComboBoxHave(modAddinConst.MENU_TOOLS)
        Case modAddinConst.TYPE_ALL_VBAPROJECT:
            For Each vbComp In Application.VBE.ActiveVBProject.VBComponents
                Call TrimLinesTabAndSpase(vbComp.CodeModule)
            Next vbComp
        Case modAddinConst.TYPE_SELECTED_MODULE:
            Call TrimLinesTabAndSpase(Application.VBE.ActiveCodePane.CodeModule)
    End Select
    Exit Sub
ErrorHandler:
    If Err.Number <> 91 Then
        Debug.Print "Error in CutTab" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "at line " & Erl
        'Call WriteErrorLog("CutTab")
    End If
    Err.Clear
End Sub

'* * * * * *
'* Sub        : ReBild - Rebuilds VBA code indentation
'* Created    : 23-03-2023 10:37
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * *
Public Sub ReBild()
    Dim moCM        As CodeModule
    Dim vbComp      As VBIDE.VBComponent
    On Error GoTo ErrorHandler
    Select Case WhatIsTextInComboBoxHave(modAddinConst.MENU_TOOLS)
        Case modAddinConst.TYPE_ALL_VBAPROJECT:
            For Each vbComp In Application.VBE.ActiveVBProject.VBComponents
                Set moCM = vbComp.CodeModule
                Call RebuildModule(moCM, moCM.Parent.Name, 1, moCM.CountOfLines, 0)
            Next vbComp
        Case modAddinConst.TYPE_SELECTED_MODULE:
            Set moCM = Application.VBE.ActiveCodePane.CodeModule
            Call RebuildModule(moCM, moCM.Parent.Name, 1, moCM.CountOfLines, 0)
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
''''''''''''
' Function:   RebuildModule
'
' Comments:   This procedure goes through the lines in a module,
'             rebuilding the code's indenting.
'
' Arguments: modCode    - The code module to indent
'             sName      - The display name of the item being indented
'             iStartLine - Value giving the line to start indenting from
'             iEndLine   - Value giving the line to end indenting at
'             iProgDone  - Value giving how much indenting has been done in total
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : RebuildModule - Rebuilds indentation in a VBA module
'* Created    : 23-03-2023 10:37
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                                     Description
'*
'* ByRef modCode As CodeModule                   : VBA Module
'* ByRef sName As String                         : Module name
'* ByRef iStartLine As Long                      : Starting line for indentation
'* ByRef iEndline As Long                        : Ending line for indentation
'* ByRef iProgDone As Long                       : Progress indicator
'* Optional ByRef mbEnableUndo As Boolean = True : Enable undo functionality
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub RebuildModule( _
        ByRef modCode As CodeModule, _
        ByRef sName As String, _
        ByRef iStartLine As Long, _
        ByRef iEndline As Long, _
        ByRef iProgDone As Long, _
        Optional ByRef mbEnableUndo As Boolean = True)
    Dim asCode() As String, asOriginal() As String, i As Long
    If iEndline = 0 Then Exit Sub    'On Error Resume Next
    ReDim asCode(0 To iEndline - iStartLine)
    ReDim asOriginal(0 To iEndline - iStartLine)
    'To save undo information? If yes, save it
    If mbEnableUndo Then
        piUndoCount = piUndoCount + 1
        'Make some space in our undo array
        If piUndoCount = 1 Then
            ReDim pauUndo(1 To 1)
        Else
            ReDim Preserve pauUndo(1 To piUndoCount)
        End If
        'Store the undo information
        With pauUndo(piUndoCount)
            Set .oMod = modCode
            .sName = sName
            .lStartLine = iStartLine
            .lEndLine = iEndline
            ReDim .asIndented(0 To iEndline - iStartLine)
            ReDim .asOriginal(0 To iEndline - iStartLine)
        End With
    End If
    'Read code module into an array and store the original code in our undo array
    For i = 0 To iEndline - iStartLine
        asCode(i) = modCode.Lines(iStartLine + i, 1)
        asOriginal(i) = asCode(i)
        If mbEnableUndo Then pauUndo(piUndoCount).asOriginal(i) = asCode(i)
    Next
    'Indent the array, showing the progress
    RebuildCodeArray asCode, sName, iProgDone
    'Copy the changed code back into the module and store in our undo array
    For i = 0 To iEndline - iStartLine
        If asOriginal(i) <> asCode(i) Then
            On Error Resume Next
            modCode.ReplaceLine iStartLine + i, asCode(i)
            On Error GoTo 0
        End If
        If mbEnableUndo Then pauUndo(piUndoCount).asIndented(i) = asCode(i)
    Next
End Sub
'* * * * * *
'* Sub        : RebuildCodeArray - Rebuilds indentation in an array
'* Created    : 23-03-2023 10:42
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):         Description
'*
'* ByRef asCodeLines( : Array of code lines
'*
'* * * * * *
Public Sub RebuildCodeArray( _
        ByRef asCodeLines() As String, _
        ByRef sName As String, _
        ByRef iProgDone As Long)
    'Variables used for the indenting code
    Dim X As Integer, i As Integer, j As Integer, k As Integer, iGap As Integer, iLineAdjust As Integer
    Dim lLineCount As Long, iCommentStart As Long, iStart As Long, iScan As Long, iDebugAdjust As Integer
    Dim iIndents As Integer, iIndentNext As Integer, iIn As Integer, iOut As Integer
    Dim iFunctionStart As Long, iParamStart As Long
    Dim bInCmt As Boolean, bProcStart As Boolean, bAlign As Boolean, bFirstCont As Boolean
    Dim bAlreadyPadded As Boolean, bFirstDim As Boolean
    Dim sLine As String, sLeft As String, sRight As String, sMatch As String, sItem As String
    Dim vaScope As Variant, vaStatic As Variant, vaType As Variant, vaInProc As Variant
    Dim iCodeLineNum As Long, sCodeLineNum As String, sOrigLine As String
    Dim OptionsTb   As ListObject
    Set OptionsTb = shSettings.ListObjects(modAddinConst.TB_OPTIONS_IDEDENT)
    On Error Resume Next
    With OptionsTb.ListColumns(2)
        mbNoIndent = False
        mbInIf = False
        'Read the indenting options from the registry
        miIndentSpaces = .Range(2, 1)    'Read VB's own setting for tab width
        mbIndentProc = .Range(3, 1)
        mbIndentFirst = .Range(4, 1)
        mbIndentDim = .Range(5, 1)
        mbIndentCmt = .Range(6, 1)
        mbIndentCase = .Range(7, 1)
        mbAlignCont = .Range(8, 1)
        mbAlignIgnoreOps = .Range(9, 1)
        mbDebugCol1 = .Range(10, 1)
        mbAlignDim = .Range(11, 1)
        miAlignDimCol = .Range(12, 1)

        mbCompilerStuffCol1 = .Range(13, 1)
        mbIndentCompilerStuff = .Range(14, 1)

        msEOLComment = .Range(15, 1)
        miEOLAlignCol = .Range(16, 1)
    End With

    If mbCompilerStuffCol1 = True Or mbIndentCompilerStuff = True Then
        mbInitialised = False
    End If

    ' Create the list of items to match for the indenting at procedure level
    If Not mbInitialised Then
        vaScope = Array(vbNullString, "Public ", "Private ", "Friend ")
        vaStatic = Array(vbNullString, "Static ")
        vaType = Array("Sub", "Function", "Property Let", "Property Get", "Property Set", "Type", "Enum")
        X = 1
        ReDim vaInProc(1)
        For i = 1 To UBound(vaScope)
            For j = 1 To UBound(vaStatic)
                For k = 1 To UBound(vaType)
                    ReDim Preserve vaInProc(X)
                    vaInProc(X) = vaScope(i) & vaStatic(j) & vaType(k)
                    X = X + 1
                Next
            Next
        Next
        ArrayFromVariant masInProc, vaInProc
        'Items to match when outdenting at procedure level
        ArrayFromVariant masOutProc, Array("End Sub", "End Function", "End Property", "End Type", "End Enum")
        If mbIndentCompilerStuff Then
            'Items to match when indenting within a procedure
            ArrayFromVariant masInCode, Array("If", "ElseIf", "Else", "#If", "#ElseIf", "#Else", "Select Case", "Case", "With", "For", "Do", "While")
            'Items to match when outdenting within a procedure
            ArrayFromVariant masOutCode, Array("ElseIf", "Else", "End If", "#ElseIf", "#Else", "#End If", "Case", "End Select", "End With", "Next", "Loop", "Wend")
        Else
            'Items to match when indenting within a procedure
            ArrayFromVariant masInCode, Array("If", "ElseIf", "Else", "Select Case", "Case", "With", "For", "Do", "While")
            'Items to match when outdenting within a procedure
            ArrayFromVariant masOutCode, Array("ElseIf", "Else", "End If", "Case", "End Select", "End With", "Next", "Loop", "Wend")
        End If
        'Items to match for declarations
        ArrayFromVariant masDeclares, Array("Dim", "Const", "Static", "Public", "Private", "#Const")
        'Things to look for within a line of code for special handling
        ArrayFromVariant masLookFor, Array("""", ": ", " As ", "'", "Rem ", "Stop ", "#If ", "#ElseIf ", "#Else ", "#End If ", "#Const ", "Debug.Print ", "Debug.Assert ")
        mbInitialised = True
    End If
    'Things to skip when finding the function start of a line
    ArrayFromVariant masFnAlign, Array("Set ", "Let ", "LSet ", "RSet ", "Declare Function", "Declare Sub", "Private Declare Function", "Private Declare Sub", "Public Declare Function", "Public Declare Sub")
    If masInCode(UBound(masInCode)) <> "Select Case" And mbIndentCase Then
        'If extra-indenting within Select Case, ensure that we have two items in the arrays
        ReDim Preserve masInCode(UBound(masInCode) + 1)
        masInCode(UBound(masInCode)) = "Select Case"
        ReDim Preserve masOutCode(UBound(masOutCode) + 1)
        masOutCode(UBound(masOutCode)) = "End Select"
    ElseIf masInCode(UBound(masInCode)) = "Select Case" And Not mbIndentCase Then
        'If not extra-indenting within Select Case, ensure that we have one item in the arrays
        ReDim Preserve masInCode(UBound(masInCode) - 1)
        ReDim Preserve masOutCode(UBound(masOutCode) - 1)
    End If
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
        If Not (mbContinued Or bInCmt) Then
            mbFirstProcLine = False
            iIndentNext = 0
            iCommentStart = 0
            iIndents = iIndents + iDebugAdjust
            iDebugAdjust = 0
            iFunctionStart = 0
            iParamStart = 0
            i = InStr(1, sLine, " ")
            If i > 0 Then
                If IsNumeric(Left$(sLine, i - 1)) Then
                    iCodeLineNum = Val(Left$(sLine, i - 1))
                    sLine = Trim$(Mid$(sLine, i + 1))
                    sOrigLine = Space(i) & Mid$(sOrigLine, i + 1)
                End If
            End If
        End If
        'Is there anything on the line?
        If Len(sLine) > 0 Then
            ' Remove leading Tabs
            Do Until Left$(sLine, 1) <> Chr$(miTAB)
                sLine = Mid$(sLine, 2)
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
            If mbContinued And mbAlignCont Then
                If mbAlignIgnoreOps And Left$(sLine, 2) = ", " Then iParamStart = iFunctionStart - 2
                If mbAlignIgnoreOps And (Mid$(sLine, 2, 1) = " " Or Left$(sLine, 2) = ":=") And Left$(sLine, 2) <> ", " Then
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
                            sLine = Left$(sLine, iScan + 1) & Trim$(Mid$(sLine, iScan + 2))
                            'And check the indenting for the line segment
                            CheckLine Mid$(sLine, iStart, iScan - 1), iIn, iOut, bProcStart
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
                        If mbAlignDim Then
                            bAlign = mbNoIndent    'Don't need to check within Type
                            If Not bAlign Then
                                ' Check if we start with a declaration item
                                For i = LBound(masDeclares) To UBound(masDeclares)
                                    sMatch = masDeclares(i) & " "
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
                                    If mbIndentProc And bFirstDim And Not mbIndentDim And Not mbNoIndent Then
                                        iGap = miAlignDimCol - Len(RTrim$(Left$(sLine, iScan)))
                                        'Adjust for a line number at the start of the line
                                        If iCodeLineNum > -1 Then iGap = iGap - Len(CStr(iCodeLineNum)) - 1
                                    Else
                                        iGap = miAlignDimCol - Len(RTrim$(Left$(sLine, iScan))) - iIndents * miIndentSpaces
                                        'Adjust for a line number at the start of the line
                                        If iCodeLineNum > -1 Then
                                            If Len(CStr(iCodeLineNum)) >= iIndents * miIndentSpaces Then
                                                iGap = iGap - (Len(CStr(iCodeLineNum)) - iIndents * miIndentSpaces) - 1
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
                                sLine = sLeft & Space$(iGap) & Mid$(sLine, iScan + 1)
                                'Update the counters
                                iLineAdjust = iLineAdjust + iGap + Len(sLeft) - iScan
                                iScan = Len(sLeft) + iGap + 3
                            End If
                        Else
                            'Not aligning Dims, so remove any existing spacing
                            iScan = Len(RTrim$(Left$(sLine, iScan)))
                            sLine = RTrim$(Left$(sLine, iScan)) & " " & Trim$(Mid$(sLine, iScan + 1))
                            iScan = iScan + 3
                        End If
                    Case "'", "Rem "
                        'The start of a comment => Handle end-of-line comments properly
                        If iScan = 1 Then
                            'New comment at start of line
                            If bProcStart And Not mbIndentFirst And Not mbNoIndent Then
                                'No indenting
                            ElseIf mbIndentCmt Or bProcStart Or mbNoIndent Then
                                'Inside the procedure, so indent to align with code
                                sLine = Space$(iIndents * miIndentSpaces) & sLine
                                iCommentStart = iScan + iIndents * miIndentSpaces
                            ElseIf iIndents > 0 And mbIndentProc And Not bProcStart Then
                                'At the top of the procedure, so indent once if required
                                sLine = Space$(miIndentSpaces) & sLine
                                iCommentStart = iScan + miIndentSpaces
                            End If
                        Else
                            'New comment at the end of a line
                            'Make sure it's a proper 'Rem'
                            If sItem = "Rem " And Mid$(sLine, iScan - 1, 1) <> " " And Mid$(sLine, iScan - 1, 1) <> ":" Then GoTo PTR_NEXT_PART
                            'Check the indenting of the previous code segment
                            CheckLine Mid$(sLine, iStart, iScan - 1), iIn, iOut, bProcStart
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
                            sRight = Trim$(Mid$(sLine, iScan))
                            'Indent the code part of the line
                            If bAlreadyPadded Then
                                sLine = RTrim$(Left$(sLine, iScan - 1))
                            Else
                                If mbContinued Then
                                    sLine = Space$((iIndents + 2) * miIndentSpaces) & sLeft
                                Else
                                    If mbIndentProc And bFirstDim And Not mbIndentDim Then
                                        sLine = sLeft
                                    Else
                                        sLine = Space$(iIndents * miIndentSpaces) & sLeft
                                    End If
                                End If
                            End If
                            mbContinued = (Right$(Trim$(sLine), 2) = " _")
                            'How do we handle end-of-line comments?
                            Select Case msEOLComment
                                Case "Absolute"
                                    iScan = iScan - iLineAdjust + Len(sOrigLine) - Len(LTrim$(sOrigLine))
                                    iGap = iScan - Len(sLine) - 1
                                Case "SameGap"
                                    iScan = iScan - iLineAdjust + Len(sOrigLine) - Len(LTrim$(sOrigLine))
                                    iGap = iScan - Len(RTrim$(Left$(sOrigLine, iScan - 1))) - 1
                                Case "StandardGap"
                                    iGap = miIndentSpaces * 2
                                Case "AlignInCol"
                                    iGap = miEOLAlignCol - Len(sLine) - 1
                            End Select
                            'Adjust for a line number at the start of the line
                            If iCodeLineNum > -1 Then
                                Select Case msEOLComment
                                    Case "Absolute", "AlignInCol"
                                        If Len(CStr(iCodeLineNum)) >= iIndents * miIndentSpaces Then
                                            iGap = iGap - (Len(CStr(iCodeLineNum)) - iIndents * miIndentSpaces) - 1
                                        End If
                                End Select
                            End If
                            If iGap < 2 Then iGap = miIndentSpaces
                            iCommentStart = Len(sLine) + iGap
                            'Put the comment in the required column
                            sLine = sLine & Space$(iGap) & sRight
                        End If
                        'Work out where the text of the comment starts, to align the next line
                        If Mid$(sLine, iCommentStart, 4) = "Rem " Then iCommentStart = iCommentStart + 3
                        If Mid$(sLine, iCommentStart, 1) = "'" Then iCommentStart = iCommentStart + 1
                        Do Until Mid$(sLine, iCommentStart, 1) <> " "
                            iCommentStart = iCommentStart + 1
                        Loop
                        iCommentStart = iCommentStart - 1
                        'Adjust for a line number at the start of the line
                        If iCodeLineNum > -1 Then
                            If Len(CStr(iCodeLineNum)) >= iIndents * miIndentSpaces Then
                                iCommentStart = iCommentStart + (Len(CStr(iCodeLineNum)) - iIndents * miIndentSpaces) + 1
                            End If
                        End If
                        'Remember if we're in a continued comment line
                        bInCmt = Right$(Trim$(sLine), 2) = " _"
                        'Rest of line is comment, so no need to check any more
                        GoTo PTR_REPLACE_LINE
                    Case "Stop ", "Debug.Print ", "Debug.Assert "
                        'A debugging statement - do we want to force to column 1?
                        If mbDebugCol1 And iStart = 1 And iScan = 1 Then
                            iLineAdjust = iLineAdjust - (Len(sOrigLine) - LTrim$(Len(sOrigLine)))
                            iDebugAdjust = iIndents
                            iIndents = 0
                        End If
                    Case "#If ", "#ElseIf ", "#Else ", "#End If ", "#Const "
                        'Do we want to force compiler directives to column 1?
                        If mbCompilerStuffCol1 And iStart = 1 And iScan = 1 Then
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
                If Not mbContinued Then bProcStart = False
                'Check the indenting of the remaining code segment
                CheckLine Mid$(sLine, iStart), iIn, iOut, bProcStart
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
            If mbFirstProcLine Then iIndents = 0
            ' What about line continuations?  Here, I indent the continued line by
            ' two indents, and check for the end of the continuations.  Note
            ' that Excel won't allow comments in the middle of line continuations
            ' and that comments are treated differently above.
            If mbContinued Then
                If Not mbAlignCont Then
                    sLine = Space$((iIndents + 2) * miIndentSpaces) & sLine
                End If
            Else
                ' Check if we start with a declaration item
                bAlign = False
                If mbIndentProc And bFirstDim And Not mbIndentDim And Not bProcStart Then
                    For i = LBound(masDeclares) To UBound(masDeclares)
                        sMatch = masDeclares(i) & " "
                        If Left$(sLine, Len(sMatch)) = sMatch Then
                            bAlign = True
                            Exit For
                        End If
                    Next
                End If
                'Not a declaration item to left-align, so pad it out
                If Not bAlign Then
                    If Not bProcStart Then bFirstDim = False
                    sLine = Space$(iIndents * miIndentSpaces) & sLine
                End If
            End If
            mbContinued = (Right$(Trim$(sLine), 2) = " _")
        End If    'Anything there?
PTR_REPLACE_LINE:
        'Add the code line number back in
        If iCodeLineNum > -1 Then
            sCodeLineNum = CStr(iCodeLineNum)
            If Len(Trim$(Left$(sLine, Len(sCodeLineNum) + 1))) = 0 Then
                sLine = sCodeLineNum & Mid$(sLine, Len(sCodeLineNum) + 1)
            Else
                sLine = sCodeLineNum & " " & Trim$(sLine)
            End If
        End If
        asCodeLines(lLineCount) = RTrim$(sLine)
        'If it's not a continued line, update the indenting for the following lines
        If Not mbContinued Then
            iIndents = iIndents + iIndentNext
            iIndentNext = 0
            If iIndents < 0 Then iIndents = 0
        Else
            'A continued line, so if we're not in a comment and we want smart continuing,
            'work out which to continue from
            If mbAlignCont And Not bInCmt Then
                If Left$(Trim$(sLine), 2) = "& " Or Left$(Trim$(sLine), 2) = "+ " Then sLine = "  " & sLine
                iFunctionStart = fnAlignFunction(sLine, bFirstCont, iParamStart)
                If iFunctionStart = 0 Then
                    iFunctionStart = (iIndents + 2) * miIndentSpaces
                    iParamStart = iFunctionStart
                End If
            End If
        End If
        bFirstCont = Not mbContinued
    Next
End Sub
'
'  Find the first occurrence of one of our key items in the list
'
'    Returns the text of the item found
'    Updates the iFrom parameter to point to the location of the found item
'
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : fnFindFirstItem - Finds the first occurrence of a key item in a line
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
    For iItem = LBound(masLookFor) To UBound(masLookFor)
        'What to find?
        sItem = masLookFor(iItem)
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
'* Function   : CheckLine - Checks a line for indenting requirements
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
    If Not mbNoIndent Then
        ' Check for indenting within the code
        For i = LBound(masInCode) To UBound(masInCode)
            sMatch = masInCode(i)
            If (Left$(sLine, Len(sMatch)) = sMatch) And ((Mid$(sLine, Len(sMatch) + 1, 1) = " ") Or (Mid$(sLine, Len(sMatch) + 1, 1) = ":")) Then
                iIndentNext = iIndentNext + 1
            End If
        Next
        ' Check for out-denting within the code
        For i = LBound(masOutCode) To UBound(masOutCode)
            sMatch = masOutCode(i)
            'Check at start of line for 'real' outdenting
            If (Left$(sLine, Len(sMatch)) = sMatch) And ((Mid$(sLine, Len(sMatch) + 1, 1) = " ") Or (Mid$(sLine, Len(sMatch) + 1, 1) = ":" And Mid$(sLine, Len(sMatch) + 2, 1) <> "=")) Then
                iOutdentThis = iOutdentThis + 1
            End If
        Next
    End If
    'Check procedure-level indenting
    For i = LBound(masInProc) To UBound(masInProc)
        sMatch = masInProc(i)
        If (Left$(sLine, Len(sMatch)) = sMatch) And ((Mid$(sLine, Len(sMatch) + 1, 1) = " ") Or (Mid$(sLine, Len(sMatch) + 1, 1) = ":" And Mid$(sLine, Len(sMatch) + 2, 1) <> "=")) Then
            bProcStart = True
            mbFirstProcLine = True
            'Don't indent within Type or Enum constructs
            If Right$(sMatch, 4) = "Type" Or Right$(sMatch, 4) = "Enum" Then
                iIndentNext = iIndentNext + 1
                mbNoIndent = True
            ElseIf mbIndentProc And Not mbNoIndent Then
                iIndentNext = iIndentNext + 1
            End If
            Exit For
        End If
    Next
    'Check procedure-level outdenting
    For i = LBound(masOutProc) To UBound(masOutProc)
        sMatch = masOutProc(i)
        If (Left$(sLine, Len(sMatch)) = sMatch) And ((Mid$(sLine, Len(sMatch) + 1, 1) = " ") Or (Mid$(sLine, Len(sMatch) + 1, 1) = ":" And Mid$(sLine, Len(sMatch) + 2, 1) <> "=")) Then
            'Don't indent within Type or Enum constructs
            If Right$(sMatch, 4) = "Type" Or Right$(sMatch, 4) = "Enum" Or mbIndentProc Then
                iOutdentThis = iOutdentThis + 1
                mbNoIndent = False
            End If
            Exit For
        End If
    Next
    'If we're not indenting, no need to consider the special cases
    If mbNoIndent Then Exit Function
    ' Treat If as a special case.  If anything other than a comment follows
    ' the Then, we don't indent
    If Left$(sLine, 3) = "If " Or Left$(sLine, 4) = "#If " Or mbInIf Then
        If mbInIf Then iIndentNext = 1
        'Strip any strings from the line
        i = InStr(1, sLine, """")
        Do Until i = 0
            j = InStr(i + 1, sLine, """")
            If j = 0 Then j = Len(sLine)
            sLine = Left$(sLine, i - 1) & Mid$(sLine, j + 1)
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
        mbInIf = (Right$(Trim$(sLine), 2) = " _")
        If i > 0 Then
            ' If there's something after the Then, we don't indent the If
            If Trim$(Mid$(sLine, i + 5)) <> vbNullString Then iIndentNext = 0
            ' No need to check next time around
            mbInIf = False
        End If
        If mbInIf Then iIndentNext = 0
    End If
End Function
' Convert a Variant array to a string array for faster comparisons
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : ArrayFromVariant - Converts a variant array to a string array for faster comparisons
'* Created    : 23-03-2023 10:45
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):     Description
'*
'* ByRef asString( :
'*
'* * * * * *
Private Sub ArrayFromVariant(ByRef asString() As String, ByRef vaVariant As Variant)
    Dim iLow As Integer, iHigh As Integer, i As Integer
    On Error Resume Next
    iLow = LBound(vaVariant)
    iHigh = UBound(vaVariant)
    ReDim asString(iLow To iHigh)
    For i = iLow To iHigh
        asString(i) = vaVariant(i)
    Next
End Sub
' Locate the start of the first parameter on the line
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : fnAlignFunction - Locates the start of the first parameter in a line
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
            sLine = Mid$(sLine, iChar + 1)
            iLPad = iChar
        End If
    End If
    iLPad = iLPad + Len(sLine) - Len(LTrim$(sLine))
    iFirstThisLine = coBrackets.Count
    sLine = Trim$(sLine)
    iCheck = 1
    'Skip over stuff that we don't want to locate the start off
    For iChar = LBound(masFnAlign) To UBound(masFnAlign)
        sMatch = masFnAlign(iChar)
        If Left$(sLine, Len(sMatch)) = sMatch Then
            iCheck = iCheck + Len(sMatch) + 1
            Exit For
        End If
    Next
    iBrackets = 0
    iSpace = 999
    For iChar = iCheck To Len(sLine)
        Select Case Mid$(sLine, iChar, 1)
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
                If Mid$(sLine, iChar, 3) = " = " Then
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
                    If Mid$(sLine, iChar - 5, 6) = " Then " Then
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
                If Mid$(sLine, iChar, 2) = ":=" Then
                    'A named paremeter => remember to align to after the name
                    vAlign(1) = ","
                    vAlign(2) = iChar + iLPad + 2
                    coBrackets.Add vAlign
                ElseIf Mid$(sLine, iChar, 2) = ": " Then
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
            iCheck = miIndentSpaces * 2 + iLPad
        Else
            iCheck = iLPad
        End If
    End If
    If iParamStart = 0 Then iParamStart = iCheck + 1
    fnAlignFunction = iCheck + 1
End Function
