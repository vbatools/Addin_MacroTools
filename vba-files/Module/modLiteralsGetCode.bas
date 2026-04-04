Attribute VB_Name = "modLiteralsGetCode"
Option Explicit
Option Private Module


Public Function parserLiteralsFormCode(ByRef wb As Workbook, bLibDeclarationDel As Boolean) As Dictionary
    Dim vbProj      As VBIDE.vbProject
    Dim VBCom       As VBIDE.vbComponent
    Dim codeMod     As VBIDE.codeModule
    Dim oDic        As Dictionary
    Dim iLineCode   As Long
    Dim sCode       As String
    Dim sLine       As String
    Dim arrLine     As Variant
    Dim iRow        As Long

    Set vbProj = wb.vbProject
    Set oDic = New Dictionary
    For Each VBCom In vbProj.VBComponents
        Set codeMod = VBCom.codeModule
        iLineCode = codeMod.CountOfLines
        If iLineCode > 0 Then
            sCode = clearCodeStrings(codeMod.Lines(1, iLineCode))
            If VBA.Len(sCode) > 0 Then
                arrLine = VBA.Split(sCode, vbNewLine)
                For iRow = LBound(arrLine) To UBound(arrLine)
                    sLine = arrLine(iRow)
                    If bLibDeclarationDel Then
                        If sLine Like "*Declare * Lib *(*)*" Then
                            GoTo dontAddDic
                        End If
                    End If
                    If VBA.Len(sLine) > 0 Then
                        If InStr(1, sLine, QUOTE_CHAR, vbBinaryCompare) > 0 Then
                            Call ExtractQuotedStrings(oDic, VBCom.Name, sLine)
                        End If
                    End If
dontAddDic:
                Next iRow
            End If
        End If
    Next VBCom
    Set parserLiteralsFormCode = oDic
End Function

Private Sub ExtractQuotedStrings(ByRef oDic As Dictionary, ByVal sNameModule As String, ByVal sLineCode As String)
    Dim currentString As String
    Dim inQuotes    As Boolean
    Dim i           As Long
    Dim char        As String
    Dim nextChar    As String
    Dim sKey        As String
    Dim arrData(1 To 1, 1 To 2) As String

    For i = 1 To Len(sLineCode)
        char = mid$(sLineCode, i, 1)
        If char = QUOTE_CHAR Then
            If inQuotes Then
                If i < Len(sLineCode) Then
                    nextChar = mid$(sLineCode, i + 1, 1)
                    If nextChar = QUOTE_CHAR Then
                        currentString = currentString & QUOTE_CHAR & QUOTE_CHAR
                        i = i + 1
                    Else
                        ' This is the closing quote of the string
                        sKey = sNameModule & "." & currentString
                        If Not oDic.Exists(sKey) And Len(currentString) > 0 Then
                            arrData(1, 1) = sNameModule
                            arrData(1, 2) = currentString
                            Call oDic.Add(sKey, arrData)
                        End If
                        currentString = vbNullString
                        inQuotes = False
                    End If
                Else
                    ' Quote at the very end of the text
                    sKey = sNameModule & "." & currentString
                    If Not oDic.Exists(sKey) And Len(currentString) > 0 Then
                        arrData(1, 1) = sNameModule
                        arrData(1, 2) = currentString
                        Call oDic.Add(sKey, arrData)
                    End If
                    inQuotes = False
                End If
            Else
                ' This is the opening quote, start recording
                inQuotes = True
            End If
        ElseIf inQuotes Then
            ' If we are inside quotes and this is a regular character, add it
            currentString = currentString & char
        End If
    Next i
End Sub