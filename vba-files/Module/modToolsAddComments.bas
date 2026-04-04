Attribute VB_Name = "modToolsAddComments"
Option Explicit
Option Private Module
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : T_AddCommentsProc - Module for auto documenting VBA project code
'* Created    : 20-01-2020 15:56
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

'abbreviated tabs
Private Const vbTab2 = vbTab & vbTab
Private Const vbTab4 = vbTab2 & vbTab2
Private Const sUPDATE As String = "'* Updated      :"
Public Const sCREATED As String = "'* TODO Created :"

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : sysAddHeaderTop - create main comment header
'* Created    : 20-01-2020 15:56
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub sysAddHeaderTop()
    Call sysAddHeader(Application.VBE.ActiveCodePane)
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : sysAddModifiedTop - create modification update line
'* Created    : 20-01-2020 15:56
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub sysAddModifiedTop()
    Call sysAddModified(Application.VBE.ActiveCodePane)
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : sysAddTODOTop - create TODO comment
'* Created    : 20-01-2020 15:56
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub sysAddTODOTop()
    Call sysAddTODO(Application.VBE.ActiveCodePane)
End Sub

Public Function addStringDelimetr() As String
    addStringDelimetr = VBA.Replace(VBA.String$(90, "*"), "**", "* ")
End Function

Public Function addArrFromTBComments() As Variant
    ReDim arr(0 To 9, 1 To 2) As String
    Dim arrVal      As Variant
    arrVal = shSettings.ListObjects(modAddinConst.TB_COMMENTS).DataBodyRange.Value2

    arr(1, 1) = arrVal(1, 2)
    If arr(1, 1) = vbNullString Then arr(1, 1) = Environ("UserName")

    arr(1, 2) = "'* Author       :" & vbTab & arr(1, 1)
    arr(2, 1) = arrVal(2, 2)
    If arr(2, 1) <> vbNullString Then arr(2, 2) = "'* Contacts     :" & vbTab & arr(2, 1)
    arr(3, 1) = arrVal(3, 2)
    If arr(3, 1) <> vbNullString Then arr(3, 2) = "'* Copyright    :" & vbTab & arr(3, 1)
    arr(4, 1) = arrVal(4, 2)
    If arr(4, 1) <> vbNullString Then arr(4, 2) = "'* Other        :" & vbTab & arr(4, 1)
    arr(5, 1) = VBA.Format$(VBA.Now, modAddinConst.FORMAT_DATE)
    arr(5, 2) = "'* Created      :" & vbTab & arr(5, 1) & vbTab

    arr(6, 1) = arr(5, 1)
    arr(6, 2) = "'* Modified     :" & vbTab & "Date and Time" & vbTab2 & "Author" & vbTab4 & "Description"
    arr(7, 2) = sUPDATE & vbTab & arr(5, 1) & vbTab & arr(1, 1) & vbTab2

    arr(8, 2) = sCREATED & arr(5, 1) & " Author: " & arr(1, 1)

    Dim i           As Byte
    For i = 1 To 5
        If arr(i, 2) <> vbNullString Then
            If arr(0, 2) <> vbNullString Then arr(0, 2) = arr(0, 2) & vbCrLf
            arr(0, 2) = arr(0, 2) & arr(i, 2)
        End If
    Next i

    addArrFromTBComments = arr
End Function

Public Function TypeProcedyreComments(ByVal sProcDeclartion As String) As String
    Dim sType       As String
    sType = TypeProcedyre(sProcDeclartion)
    TypeProcedyreComments = "* " & sType & VBA.Space$(15 - VBA.Len("* " & sType)) & ":"
End Function

Public Function TypeModuleComments(ByVal sTypeModule As vbext_ComponentType) As String
    Select Case sTypeModule
             Case vbext_ct_StdModule
            TypeModuleComments = "* Module       :"
        Case vbext_ct_ClassModule
            TypeModuleComments = "* Class        :"
        Case vbext_ct_MSForm
            TypeModuleComments = "* UserForm     :"
        Case vbext_ct_Document
            TypeModuleComments = "* Document     :"
        Case Else
            TypeModuleComments = "* Module       :"
    End Select
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : sysAddHeader - create main comment header
'* Created    : 20-01-2020 15:56
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                         Description
'*
'* ByRef CurentCodePane As CodePane : active VBE code pane
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub sysAddHeader(ByRef CurentCodePane As CodePane)
    Dim nLine       As Long
    Dim i           As Byte
    Dim procKind    As VBIDE.vbext_ProcKind
    Dim sProc       As String
    Dim sTemp       As String
    Dim sType       As String
    Dim sProcDeclartion As String
    Dim sProcArguments As String
    Dim sComments   As String
    sComments = addArrFromTBComments()(0, 2)

    On Error Resume Next
    With CurentCodePane
        'get start line and name of current procedure
        sProc = GetCurrentProcInfo(nLine, CurentCodePane)

        'create '* * *' block separator line
        sTemp = addStringDelimetr()

        'setup a type label
        If sProc = "" Then
            'top of module
            sProc = .codeModule.Name
            sType = TypeModuleComments(.codeModule.Parent.Type)
            nLine = 1
        Else
            For i = 0 To 4
                procKind = i
                sProcDeclartion = GetProcedureDeclaration(.codeModule, sProc, procKind, LineSplitRemove)
                If sProcDeclartion <> vbNullString Then Exit For
            Next
            sProcArguments = AddStringParamertFromProcedureDeclaration(sProcDeclartion)
            sType = TypeProcedyreComments(sProcDeclartion)
        End If

        'create text block for insertion
        sTemp = "'" & sTemp & vbCrLf & _
                "'" & sType & vbTab & sProc & "- add description!" & vbCrLf & _
                sComments & vbCrLf & _
                sProcArguments & _
                "'" & sTemp
        'insert
        .codeModule.InsertLines nLine, sTemp
    End With
End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : sysAddModified - create modification update line
'* Created    : 20-01-2020 15:56
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                         Description
'*
'* ByRef CurentCodePane As CodePane : active VBE code pane
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub sysAddModified(ByRef CurentCodePane As CodePane)
    Dim nLine       As Long
    Dim sProc       As String
    Dim sSecondLine As String
    Dim arrComments As Variant
    Dim sFersLine   As String
    arrComments = addArrFromTBComments()

    On Error Resume Next
    With CurentCodePane
        'get start line and name of current procedure
        sProc = GetCurrentProcInfo(nLine, CurentCodePane)
        sFersLine = arrComments(6, 2) & vbCrLf
        sSecondLine = arrComments(7, 2)
        If Not .codeModule.Lines(nLine - 2, 1) Like sUPDATE & "*" Then
            sSecondLine = sFersLine & sSecondLine
        End If
        'insert
        .codeModule.InsertLines nLine - 1, sSecondLine
    End With
End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : sysAddTODO - create TODO comment line
'* Created    : 20-01-2020 15:56
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                         Description
'*
'* ByRef CurentCodePane As CodePane : active VBE code pane
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub sysAddTODO(ByRef CurentCodePane As CodePane)
    Dim lStartLine  As Long
    Dim lStartColumn As Long
    Dim lEndLine    As Long
    Dim lEndColumn  As Long
    Dim sFersLine   As String
    Dim sSpec       As String

    On Error Resume Next
    With CurentCodePane
        'insert
        .GetSelection lStartLine, lStartColumn, lEndLine, lEndColumn
        sSpec = VBA.String$(lStartColumn - 1, " ")
        sFersLine = sSpec & addArrFromTBComments()(8, 2) & vbCrLf & sSpec & "'*"
        .codeModule.InsertLines lStartLine, sFersLine
    End With
End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : GetCurrentProcInfo - get line number and procedure name
'* Created    : 20-01-2020 15:56
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                         Description
'*
'* ByRef nLine As Long              : line number
'* ByRef CurentCodePane As CodePane : active VBE code pane
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function GetCurrentProcInfo(ByRef nLine As Long, ByRef CurentCodePane As CodePane) As String
    Dim t           As Long

    With CurentCodePane
        'get procedure name from cursor position
        Call .GetSelection(nLine, t, t, t)
        GetCurrentProcInfo = .codeModule.ProcOfLine(nLine, vbext_pk_Proc)

        If GetCurrentProcInfo = vbNullString Then
            'we are in the declaration section; skip existing user comment lines
            Do While .codeModule.Find("'*", nLine, 1, .codeModule.CountOfDeclarationLines, 2)
                nLine = nLine + 1
                If nLine > .codeModule.CountOfDeclarationLines Then Exit Do
            Loop
        Else
            'inside procedure -> get first line number
            nLine = .codeModule.ProcBodyLine(GetCurrentProcInfo, vbext_pk_Proc)
        End If
    End With

End Function
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : AddStringParamertFromProcedureDeclaration - returns comment line with function or procedure parameters
'* Created    : 20-01-2020 15:56
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                     Description
'*
'* ByVal sPocDeclartion As String : function or procedure declaration line
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function AddStringParamertFromProcedureDeclaration(ByVal sPocDeclartion As String) As String
    Dim sDeclaration As String
    sDeclaration = Right$(sPocDeclartion, Len(sPocDeclartion) - InStr(1, sPocDeclartion, "("))
    sDeclaration = Left$(sDeclaration, InStr(1, sDeclaration, ")") - 1)
    'if no parameters then return empty
    If sDeclaration = vbNullString Then Exit Function

    Dim arStr()     As String
    Dim sTemp       As String
    Dim i           As Byte
    Dim iMaxLen     As Byte
    Dim iTempLen    As Byte

    arStr = Split(sDeclaration, ",")
    iMaxLen = 0
    For i = 0 To UBound(arStr)
        iTempLen = Len(Trim$(arStr(i)))
        If iMaxLen < iTempLen Then iMaxLen = iTempLen
    Next i

    sDeclaration = "'* Argument(s)  :" & String$(iMaxLen - Len(Trim$("'* Argument(s)  :")), " ") & vbTab2 & "Description" & vbCrLf & "'*" & vbCrLf
    For i = 0 To UBound(arStr)
        sTemp = "'* " & Trim$(arStr(i)) & String$(iMaxLen - Len(Trim$(arStr(i))), " ") & " :"
        sDeclaration = sDeclaration & sTemp & vbCrLf
    Next i
    AddStringParamertFromProcedureDeclaration = sDeclaration & "'* " & vbCrLf
End Function