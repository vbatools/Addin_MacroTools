Attribute VB_Name = "modToolsSwapEgual"
Option Explicit
Option Private Module

Public Sub SwapEgual()
    Dim NewText     As String
    Dim lineText    As String
    Dim sL          As Long
    Dim eL          As Long
    Dim sC          As Long
    Dim eC          As Long
    Dim i           As Long

    Call Application.VBE.ActiveCodePane.GetSelection(sL, sC, eL, eC)
    If sL = eL Then
        lineText = Application.VBE.ActiveCodePane.codeModule.Lines(sL, 1)
        NewText = VBA.mid(lineText, 1, sC - 1) & SwapEgualText(VBA.mid(lineText, sC, eC - sC)) & VBA.mid(lineText, eC)
        If NewText <> vbNullString Then Call Application.VBE.ActiveCodePane.codeModule.ReplaceLine(sL, NewText)
    Else
        For i = sL To eL
            NewText = vbNullString
            lineText = Application.VBE.ActiveCodePane.codeModule.Lines(i, 1)
            If i = sL Then
                NewText = VBA.mid(lineText, 1, sC - 1) & SwapEgualText(VBA.mid(lineText, sC))
            ElseIf i = eL Then
                NewText = SwapEgualText(VBA.mid(lineText, 1, eC - 1)) & VBA.mid(lineText, eC)
            Else
                NewText = SwapEgualText(lineText)
            End If
            If NewText <> vbNullString Then Call Application.VBE.ActiveCodePane.codeModule.ReplaceLine(i, NewText)
        Next i
    End If
    Call Application.VBE.ActiveCodePane.SetSelection(sL, sC, eL, eC)
End Sub

Private Function SwapEgualText(ByVal sText As String) As String
    Dim arrStr      As Variant
    Dim i           As Integer
    Dim nPos        As Long
    Dim sLine       As String
    Dim sLeft       As String
    Dim sRight      As String
    Dim sNew        As String

    arrStr = VBA.Split(sText, vbNewLine)
    For i = 0 To UBound(arrStr)
        sLine = arrStr(i)
        nPos = VBA.InStr(sLine, " = ")
        If nPos > 0 Then
            sLeft = VBA.RTrim$(Left(sLine, nPos - 1))
            sRight = VBA.Right$(sLine, Len(sLine) - nPos - 2)
            If sNew <> vbNullString Then sNew = sNew & vbNewLine
            sNew = sNew & VBA.Space(Len(sLeft) - VBA.Len(LTrim(sLeft))) & VBA.Trim$(sRight) & " = " & VBA.Trim$(sLeft)
        End If
    Next i
    SwapEgualText = sNew
End Function
