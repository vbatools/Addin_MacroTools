Attribute VB_Name = "modToolsDimOneLine"
Option Explicit
Option Private Module

Public Sub dimMultiLine()
    Call RemoveLineNumbersVBProject
    Dim objVBCitem  As VBIDE.codeModule

    Dim sCode       As String
    Dim sCodeNew    As String
    Set objVBCitem = Application.VBE.ActiveCodePane.codeModule
    sCode = GetCodeFromModule(objVBCitem)
    If sCode = vbNullString Then Exit Sub
    sCodeNew = FormatCodeToMultilineComma(sCode)
    Call SetCodeInModule(objVBCitem, sCodeNew)
    Call ReBild
End Sub

Public Sub dimOneLine()
    Call RemoveLineNumbersVBProject
    Dim objVBCitem  As VBIDE.codeModule
    Dim sCode       As String
    Dim sCodeNew    As String
    Dim sDelimetr   As String
    Set objVBCitem = Application.VBE.ActiveCodePane.codeModule
    sCode = GetCodeFromModule(objVBCitem)
    If sCode = vbNullString Then Exit Sub
    Dim arrCodsStr  As Variant
    arrCodsStr = VBA.Split(sCode, vbNewLine)
    Dim i           As Long
    Dim iCount      As Long
    iCount = UBound(arrCodsStr, 1)
    For i = iCount To 0 Step -1
        sDelimetr = vbNewLine
        If i <> 0 Then
            If Not VBA.Trim$(arrCodsStr(i)) Like "Dim *: *" And Not VBA.Trim$(arrCodsStr(i - 1)) Like "Dim *: *" Then
                If VBA.Left$(VBA.Trim$(arrCodsStr(i)), 4) = "Dim " And VBA.Left$(VBA.Trim$(arrCodsStr(i - 1)), 4) = "Dim " Then
                    sDelimetr = vbNullString
                    arrCodsStr(i) = VBA.Replace(arrCodsStr(i), "Dim ", ", ")
                    arrCodsStr(i) = WorksheetFunction.Trim(arrCodsStr(i))
                End If
            End If
        Else
            sDelimetr = vbNullString
        End If
        sCodeNew = sDelimetr & arrCodsStr(i) & sCodeNew
    Next i
    Call SetCodeInModule(objVBCitem, sCodeNew)
    Call ReBild
End Sub
