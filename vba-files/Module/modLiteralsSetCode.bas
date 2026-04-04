Attribute VB_Name = "modLiteralsSetCode"
Option Explicit
Option Private Module

Const QUOTE_CHAR    As String = """"

Public Function renameLiteralsToCode(ByRef vbProj As VBIDE.vbProject, ByRef arrCode As Variant) As Boolean

    If IsEmpty(arrCode) Then Exit Function

    Dim VBModule    As VBIDE.vbComponent
    Dim sModuleName As String
    Dim sCode       As String
    Dim i           As Long
    Dim iCount      As Long

    iCount = UBound(arrCode, 1)
    sModuleName = vbNullString
    For i = 1 To iCount
        If VBA.Len(arrCode(i, 1)) > 0 Then
            If sModuleName <> arrCode(i, 1) Then
                Call SaveModuleCode(VBModule, sCode)

                sModuleName = arrCode(i, 1)
                Set VBModule = getVBModuleByName(vbProj, sModuleName)
                sCode = vbNullString

                If Not VBModule Is Nothing Then
                    With VBModule.codeModule
                        If .CountOfLines > 0 Then
                            sCode = .Lines(1, .CountOfLines)
                        End If
                    End With
                Else
                    arrCode(i, 4) = "module not found"
                End If
            End If
            If VBA.Len(sCode) > 0 Then
                sCode = VBA.Replace(sCode, QUOTE_CHAR & arrCode(i, 2) & QUOTE_CHAR, QUOTE_CHAR & arrCode(i, 3) & QUOTE_CHAR)
                arrCode(i, 4) = "modified"
            End If
        End If
    Next i
    Call SaveModuleCode(VBModule, sCode)
    renameLiteralsToCode = True
End Function

Private Sub SaveModuleCode(ByRef VBModuleComp As VBIDE.vbComponent, ByVal sNewCode As String)
    If Not VBModuleComp Is Nothing And VBA.Len(sNewCode) > 0 Then
        With VBModuleComp.codeModule
            .DeleteLines 1, .CountOfLines
            .InsertLines 1, sNewCode
        End With
    End If
End Sub