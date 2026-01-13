Attribute VB_Name = "modToolsSwapEgual"
Option Explicit
Option Private Module
Option Compare Text
Option Base 1


Public Sub SwapEgual()
    Dim newText     As String
    Dim lineText    As String
    Dim sL          As Long
    Dim eL          As Long
    Dim sC          As Long
    Dim eC          As Long
    Dim i           As Long

    Call Application.VBE.ActiveCodePane.GetSelection(sL, sC, eL, eC)
    If sL = eL Then
        lineText = Application.VBE.ActiveCodePane.CodeModule.Lines(sL, 1)
        newText = VBA.Mid(lineText, 1, sC - 1) & SwapEgualText(VBA.Mid(lineText, sC, eC - sC)) & VBA.Mid(lineText, eC)
        If newText <> vbNullString Then Call Application.VBE.ActiveCodePane.CodeModule.ReplaceLine(sL, newText)
    Else
        For i = sL To eL
            newText = vbNullString
            lineText = Application.VBE.ActiveCodePane.CodeModule.Lines(i, 1)
            If i = sL Then
                newText = VBA.Mid(lineText, 1, sC - 1) & SwapEgualText(VBA.Mid(lineText, sC))
            ElseIf i = eL Then
                newText = SwapEgualText(VBA.Mid(lineText, 1, eC - 1)) & VBA.Mid(lineText, eC)
            Else
                newText = SwapEgualText(lineText)
            End If
            If newText <> vbNullString Then Call Application.VBE.ActiveCodePane.CodeModule.ReplaceLine(i, newText)
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
