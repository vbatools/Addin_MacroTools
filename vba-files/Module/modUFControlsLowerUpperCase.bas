Attribute VB_Name = "modUFControlsLowerUpperCase"
Option Explicit
Option Private Module
Option Compare Text
Option Base 1




'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : UperTextInControl - Changes text in controls to uppercase
'* Created    : 01-07-2022 11:12
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub UperTextInControl()
    Call LowerAndUperTextInControl(True)
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : LowerTextInControl - Changes text in controls to lowercase
'* Created    : 22-03-2023 16:07
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub LowerTextInControl()
    Call LowerAndUperTextInControl(False)
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : LowerAndUperTextInControl - Changes text in controls to upper or lower case
'* Created    : 22-03-2023 16:08
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):             Description
'*
'* ByVal bUCase As Boolean :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub LowerAndUperTextInControl(ByVal bUCase As Boolean)
    If Application.VBE.ActiveWindow.Type = vbext_wt_Designer Then
        Dim objActiveModule As VBComponent
        Set objActiveModule = getActiveModule()
        If Not objActiveModule Is Nothing Then
            Dim ctl As Control
            On Error Resume Next
            For Each ctl In objActiveModule.Designer.Selected
                If bUCase Then
                    Call CallByName(ctl, "Caption", VbLet, VBA.UCase$(CallByName(ctl, "Caption", VbGet)))
                Else
                    Call CallByName(ctl, "Caption", VbLet, VBA.LCase$(CallByName(ctl, "Caption", VbGet)))
                End If
            Next ctl
            On Error GoTo 0
        End If
    End If
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : UperTextInForm - Changes text in form to uppercase
'* Created    : 22-03-2023 16:05
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub UperTextInForm()
    Call LowerAndUperTextInForm(True)
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : LowerTextInForm - Changes text in form to lowercase
'* Created    : 22-03-2023 16:05
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub LowerTextInForm()
    Call LowerAndUperTextInForm(False)
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : LowerAndUperTextInForm - Changes text in form to upper or lower case
'* Created    : 22-03-2023 16:03
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):             Description
'*
'* ByVal bUCase As Boolean :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub LowerAndUperTextInForm(ByVal bUCase As Boolean)
    Dim oVBComp     As VBIDE.VBComponent
    Set oVBComp = Application.VBE.SelectedVBComponent
    With oVBComp
        If .Type = vbext_ct_MSForm Then
            If bUCase Then
                .Properties("Caption") = VBA.UCase$(.Properties("Caption"))
            Else
                .Properties("Caption") = VBA.LCase$(.Properties("Caption"))
            End If
        End If
    End With
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : toUpperCase - Converts selected text to uppercase
'* Created    : 18-02-2020 09:05
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub toUpperCase()
    On Error GoTo ErrorHandler
    Dim i           As Long
    Dim newText     As String
    Dim lineText    As String
    Dim sL          As Long
    Dim eL          As Long
    Dim sC          As Long
    Dim eC          As Long

    Call Application.VBE.ActiveCodePane.GetSelection(sL, sC, eL, eC)

    If sL = eL Then
        lineText = Application.VBE.ActiveCodePane.CodeModule.Lines(sL, 1)
        newText = VBA.Mid(lineText, 1, sC - 1) & VBA.UCase$(VBA.Mid(lineText, sC, eC - sC)) & VBA.Mid(lineText, eC)
        If newText <> vbNullString Then Call Application.VBE.ActiveCodePane.CodeModule.ReplaceLine(sL, newText)
    Else
        For i = sL To eL
            newText = ""
            lineText = Application.VBE.ActiveCodePane.CodeModule.Lines(i, 1)
            If i = sL Then
                newText = VBA.Mid(lineText, 1, sC - 1) & VBA.UCase$(VBA.Mid(lineText, sC))
            ElseIf i = eL Then
                newText = VBA.UCase$(VBA.Mid(lineText, 1, eC - 1)) & VBA.Mid(lineText, eC)
            Else
                newText = VBA.UCase$(lineText)
            End If
            If newText <> vbNullString Then Call Application.VBE.ActiveCodePane.CodeModule.ReplaceLine(i, newText)
        Next i
    End If
    Call Application.VBE.ActiveCodePane.SetSelection(sL, sC, eL, eC)
ErrorHandler:
    Select Case Err.Number
        Case 0
        Case Else
            Debug.Print "Error in U_UpperAndLowerCase.toUpperCase" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "at line " & Erl
            'Call WriteErrorLog("U_UpperAndLowerCase.toUpperCase")
    End Select
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : toLowerCase - Converts selected text to lowercase
'* Created    : 18-02-2020 09:06
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub toLowerCase()
    On Error GoTo ErrorHandler
    Dim i           As Long
    Dim newText     As String
    Dim lineText    As String
    Dim sL          As Long
    Dim eL          As Long
    Dim sC          As Long
    Dim eC          As Long

    Call Application.VBE.ActiveCodePane.GetSelection(sL, sC, eL, eC)

    If sL = eL Then
        lineText = Application.VBE.ActiveCodePane.CodeModule.Lines(sL, 1)
        newText = VBA.Mid(lineText, 1, sC - 1) & VBA.LCase(VBA.Mid(lineText, sC, eC - sC)) & VBA.Mid(lineText, eC)
        If newText <> vbNullString Then Call Application.VBE.ActiveCodePane.CodeModule.ReplaceLine(sL, newText)
    Else
        For i = sL To eL
            newText = ""
            lineText = Application.VBE.ActiveCodePane.CodeModule.Lines(i, 1)
            If i = sL Then
                newText = VBA.Mid(lineText, 1, sC - 1) & VBA.LCase(VBA.Mid(lineText, sC))
            ElseIf i = eL Then
                newText = VBA.LCase(VBA.Mid(lineText, 1, eC - 1)) & VBA.Mid(lineText, eC)
            Else
                newText = VBA.LCase(lineText)
            End If
            If newText <> vbNullString Then Call Application.VBE.ActiveCodePane.CodeModule.ReplaceLine(i, newText)
        Next i
    End If

    Call Application.VBE.ActiveCodePane.SetSelection(sL, sC, eL, eC)

ErrorHandler:
    Select Case Err.Number
        Case 0
        Case Else
            Debug.Print "Error in U_UpperAndLowerCase.toLowerCase" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "at line " & Erl
            'Call WriteErrorLog("U_UpperAndLowerCase.toLowerCase")
    End Select
End Sub

