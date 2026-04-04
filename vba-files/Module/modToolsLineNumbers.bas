Attribute VB_Name = "modToolsLineNumbers"
Option Explicit
Option Private Module

Public Sub RemoveLineNumbersVBProject()
      On Error GoTo ErrorHandler
      Dim VBComp      As VBIDE.vbComponent
      Dim iSelectedRow As Long
      Dim iSelectedCol As Long
      Call Application.VBE.ActiveCodePane.GetSelection(iSelectedRow, iSelectedCol, 1, 1)

    Select Case WhatIsTextInComboBoxHave(modAddinConst.MENU_TOOLS)
        Case modAddinConst.TYPE_ALL_VBAPROJECT:
            For Each VBComp In Application.VBE.ActiveVBProject.VBComponents
                Call RemoveLineNumbersModule(VBComp)
            Next VBComp
        Case modAddinConst.TYPE_SELECTED_MODULE:
            Call RemoveLineNumbersModule(Application.VBE.ActiveCodePane.codeModule.Parent)
    End Select
    If iSelectedRow = 0 Then iSelectedRow = iSelectedRow + 1
    Call Application.VBE.ActiveCodePane.SetSelection(iSelectedRow, iSelectedCol, iSelectedRow, iSelectedCol)
    Exit Sub
ErrorHandler:
    Select Case Err.Number
        Case 91:
            Exit Sub
        Case Else:
            Call WriteErrorLog("RemoveLineNumbersVBProject", False)
    End Select
    Err.Clear
End Sub

Public Sub AddLineNumbersVBProject()
    On Error GoTo ErrorHandler
    Dim VBComp      As VBIDE.vbComponent

    Dim iSelectedRow As Long
    Dim iSelectedCol As Long
    Call Application.VBE.ActiveCodePane.GetSelection(iSelectedRow, iSelectedCol, 1, 1)


    Select Case WhatIsTextInComboBoxHave(modAddinConst.MENU_TOOLS)
        Case modAddinConst.TYPE_ALL_VBAPROJECT:
            For Each VBComp In Application.VBE.ActiveVBProject.VBComponents
                Call AddLineNumbersModule(VBComp)
            Next VBComp
        Case modAddinConst.TYPE_SELECTED_MODULE:
            Call AddLineNumbersModule(Application.VBE.ActiveCodePane.codeModule.Parent)
    End Select
    If iSelectedRow = 0 Then iSelectedRow = iSelectedRow + 1
    Call Application.VBE.ActiveCodePane.SetSelection(iSelectedRow, iSelectedCol, iSelectedRow, iSelectedCol)
    Exit Sub
ErrorHandler:
    Select Case Err.Number
        Case 91:
            Exit Sub
        Case Else:
            Call WriteErrorLog("AddLineNumbersVBProject", False)
    End Select
    Err.Clear
End Sub

Public Sub AddLineNumbersModule(ByRef VBComp As VBIDE.vbComponent)
    Dim lLineIdx    As Long
    Dim lTotalLines As Long
    Dim sCurrentLine As String
    Dim sPreviousLine As String
    Dim lPaddingLength As Long
    Dim arrCode()   As String
    Dim bInProcedure As Boolean

    On Error GoTo ErrHandler

    With VBComp.codeModule
        lTotalLines = .CountOfLines
        If lTotalLines = 0 Then Exit Sub

        arrCode = VBA.Split(.Lines(1, lTotalLines), vbNewLine)
        lPaddingLength = VBA.Len(lTotalLines)

        For lLineIdx = lTotalLines To .CountOfDeclarationLines Step -1
            sCurrentLine = arrCode(lLineIdx - 1)

            sPreviousLine = vbNullString
            If lLineIdx > 1 Then sPreviousLine = arrCode(lLineIdx - 2)

            If IsProcEndLine(sCurrentLine) Then bInProcedure = True
            If IsProcStartLine(sCurrentLine) Then bInProcedure = False

            If bInProcedure Then
                If Not IsProcEndLine(sCurrentLine) Then
                    If VBA.Right$(sCurrentLine, 1) <> ":" And Not IsMultiLineString(sPreviousLine) Then

                        If IsSelectCase(sPreviousLine) Then
                            arrCode(lLineIdx - 1) = "  " & VBA.String$(lPaddingLength, " ") & arrCode(lLineIdx - 1)
                        Else

                            arrCode(lLineIdx - 1) = CStr(lLineIdx) & ": " & _
                                    VBA.String$(lPaddingLength - VBA.Len(CStr(lLineIdx)), " ") & _
                                    RemoveLineNumbers(arrCode(lLineIdx - 1), lPaddingLength)
                        End If
                    End If
                End If
            End If
        Next lLineIdx
        Call .DeleteLines(1, lTotalLines)
        Call .InsertLines(1, VBA.Join(arrCode, vbNewLine))
    End With
    Exit Sub

ErrHandler:
    Debug.Print ">> Error adding line numbers in module '" & VBComp.Name & "':" & vbCrLf & _
            "Number: " & Err.Number & vbCrLf & _
            "Description: " & Err.Description, vbCritical
End Sub

Private Sub RemoveLineNumbersModule(ByRef VBComp As VBIDE.vbComponent)
    Dim lLineIdx    As Long
    Dim lTotalLines As Long
    Dim lPaddingLength As Long
    Dim sPreviousLine As String
    Dim lrightPart  As Long
    Dim arrCode()   As String

    With VBComp.codeModule
        lTotalLines = .CountOfLines
        lPaddingLength = VBA.Len(lTotalLines)
        If lTotalLines = 0 Then Exit Sub

        arrCode = VBA.Split(.Lines(1, lTotalLines), vbNewLine)
        lPaddingLength = VBA.Len(lTotalLines)

        For lLineIdx = lTotalLines To .CountOfDeclarationLines Step -1

            sPreviousLine = vbNullString
            If lLineIdx > 1 Then
                sPreviousLine = arrCode(lLineIdx - 2)
                lrightPart = VBA.Len(arrCode(lLineIdx - 1)) - lPaddingLength - 2

                If IsSelectCase(sPreviousLine) And lrightPart > 0 Then
                    If VBA.Replace(VBA.Left$(arrCode(lLineIdx - 1), lrightPart), " ", vbNullString) = vbNullString Then
                        arrCode(lLineIdx - 1) = VBA.Right$(arrCode(lLineIdx - 1), lrightPart)
                    End If
                Else
                    arrCode(lLineIdx - 1) = RemoveLineNumbers(arrCode(lLineIdx - 1), lPaddingLength)
                End If
            End If
        Next lLineIdx
        Call .DeleteLines(1, lTotalLines)
        Call .InsertLines(1, VBA.Join(arrCode, vbNewLine))
    End With
End Sub

'================================================================================
' Helper Functions
'================================================================================
Private Function RemoveLineNumbers(ByVal sLine As String, ByRef lPaddingLength As Long) As String
    RemoveLineNumbers = sLine
    If VBA.Len(sLine) = 0 Then Exit Function

    Dim lColonPos   As Long
    Dim lDel        As Long
    Dim sPrefix     As String
    lColonPos = VBA.InStr(1, sLine, ":")

    If lColonPos > 0 Then
        ' Extract text before colon and remove possible spaces
        sPrefix = VBA.Trim$(VBA.Left$(sLine, lColonPos - 1))

        ' Check if prefix is a number (protection from removing colons in code)
        If VBA.IsNumeric(sPrefix) Then
            ' Return the part of the string after the colon.
            sLine = VBA.mid$(sLine, lColonPos + 2)
            lDel = lPaddingLength - VBA.Len(sPrefix)
            If VBA.Len(sLine) > lDel And lDel > 0 Then
                If VBA.mid(sLine, 1, lDel) = " " Then sLine = VBA.Right$(sLine, VBA.Len(sLine) - lDel)
            End If
            RemoveLineNumbers = sLine
        End If
    End If
End Function

Private Function IsSelectCase(ByRef sLineUp As String) As Boolean
    ' Checks if the previous line was a Select Case statement
         If VBA.Len(sLineUp) = 0 Then Exit Function
    If VBA.Trim$(sLineUp) Like "*Select Case *" Then IsSelectCase = True
End Function

Private Function IsMultiLineString(ByRef sLineUp As String) As Boolean
    ' Checks if the line continues to the next line (using underscore)
    If VBA.Len(sLineUp) < 2 Then Exit Function
    If VBA.Right$(sLineUp, 2) = " _" Then IsMultiLineString = True
End Function

Private Function IsProcEndLine(ByRef sLine As String) As Boolean
    Dim sCleanLine  As String
    sCleanLine = VBA.Trim$(sLine)

    If VBA.Len(sCleanLine) = 0 Then Exit Function
    If VBA.Left$(sCleanLine, 1) = "'" Then Exit Function

    IsProcEndLine = (sCleanLine Like "End Sub*") Or _
            (sCleanLine Like "End Function*") Or _
            (sCleanLine Like "End Property*")
End Function

Private Function IsProcStartLine(ByRef sLine As String) As Boolean
    If VBA.Len(sLine) = 0 Then Exit Function
    If VBA.Trim$(sLine) Like "'*" Then Exit Function

    Dim vKeywords   As Variant
    Dim i           As Long
    Dim sPattern    As String

    ' Array of procedure declaration keywords
    vKeywords = Array("Function", "Sub", "Property Get", "Property Let", "Property Set")

    For i = LBound(vKeywords) To UBound(vKeywords)
        sPattern = "*" & vKeywords(i) & " *(*"
        ' Matches "Sub MyProc(", "Function MyProc(", or line breaks like " _"
        If sLine Like sPattern & "*)*" Or sLine Like sPattern & " _" Then
            IsProcStartLine = True
            Exit Function
        End If
    Next i
End Function