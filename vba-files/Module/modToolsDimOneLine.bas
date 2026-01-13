Attribute VB_Name = "modToolsDimOneLine"
Option Explicit
Option Private Module
Option Compare Text
Option Base 1


'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : dimMultiLine - Converts single-line Dim declarations to multiple lines
'* Created    : 22-03-2023 14:25
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub dimMultiLine()
    Call RemoveLineNumbersPublic
    Dim objVBCitem  As VBIDE.CodeModule

    Dim sCode       As String
    Dim sCodeNew    As String
    Set objVBCitem = Application.VBE.ActiveCodePane.CodeModule
    sCode = GetCodeFromModule(objVBCitem)
    If sCode = vbNullString Then Exit Sub
    Dim arrCodsStr  As Variant
    arrCodsStr = VBA.Split(sCode, vbNewLine)
    Dim i           As Long
    Dim iCount      As Long
    iCount = UBound(arrCodsStr, 1)
    For i = 0 To iCount
        If Not VBA.Trim$(arrCodsStr(i)) Like "Dim *: *" And Not VBA.Trim$(arrCodsStr(i)) Like "Dim *(* To *, *" Then
            If VBA.Trim$(arrCodsStr(i)) Like "Dim *, *" Then
                arrCodsStr(i) = VBA.Replace(arrCodsStr(i), ", ", vbNewLine & "Dim ")
            End If
        End If
        If sCodeNew <> vbNullString Then sCodeNew = sCodeNew & vbNewLine
        sCodeNew = sCodeNew & arrCodsStr(i)
    Next i
    Call SetCodeInModule(objVBCitem, sCodeNew)
    Call ReBild
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : dimOneLine - Converts multiple-line Dim declarations to a single line
'* Created    : 22-03-2023 14:25
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub dimOneLine()
    Call RemoveLineNumbersPublic
    Dim objVBCitem  As VBIDE.CodeModule
    Dim sCode       As String
    Dim sCodeNew    As String
    Dim sDelimetr   As String
    Set objVBCitem = Application.VBE.ActiveCodePane.CodeModule
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
