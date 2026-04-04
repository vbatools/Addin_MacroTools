Attribute VB_Name = "modUFControlsLowerUpperCase"
Option Explicit
Option Private Module

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : UperTextInControl - change control caption/text to uppercase
'* Created    : 01-07-2022 11:12
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub UperTextInControl()
    Call LowerAndUperTextInControl(True)
End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : LowerTextInControl - change control caption/text to lowercase
'* Created    : 22-03-2023 16:07
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub LowerTextInControl()
    Call LowerAndUperTextInControl(False)
End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : LowerAndUperTextInControl - change control caption/text case to uppercase or lowercase
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
        Dim objActiveModule As vbComponent
        Set objActiveModule = getActiveModule()
        If Not objActiveModule Is Nothing Then
            Dim ctl As control
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
'* Sub        : UperTextInForm - change form text to uppercase
'* Created    : 22-03-2023 16:05
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub UperTextInForm()
    Call LowerAndUperTextInForm(True)
End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : LowerTextInForm - change form text to lowercase
'* Created    : 22-03-2023 16:05
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub LowerTextInForm()
    Call LowerAndUperTextInForm(False)
End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : LowerAndUperTextInForm - change form text to uppercase or lowercase
'* Created    : 22-03-2023 16:03
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):             Description
'*
'* ByVal bUCase As Boolean :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub LowerAndUperTextInForm(ByVal bUCase As Boolean)
    Dim VBComp      As VBIDE.vbComponent
    Set VBComp = Application.VBE.SelectedVBComponent
    With VBComp
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
'* Sub        : toUpperCase - convert selected code to UPPERCASE
'* Created    : 18-02-2020 09:05
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub toUpperCase()
    On Error GoTo ErrorHandler
    Dim i           As Long
    Dim NewText     As String
    Dim lineText    As String
    Dim sL          As Long
    Dim eL          As Long
    Dim sC          As Long
    Dim eC          As Long

    Call Application.VBE.ActiveCodePane.GetSelection(sL, sC, eL, eC)

    If sL = eL Then
        lineText = Application.VBE.ActiveCodePane.codeModule.Lines(sL, 1)
        NewText = VBA.mid(lineText, 1, sC - 1) & VBA.UCase$(VBA.mid(lineText, sC, eC - sC)) & VBA.mid(lineText, eC)
        If NewText <> vbNullString Then Call Application.VBE.ActiveCodePane.codeModule.ReplaceLine(sL, NewText)
    Else
        For i = sL To eL
            NewText = ""
            lineText = Application.VBE.ActiveCodePane.codeModule.Lines(i, 1)
            If i = sL Then
                NewText = VBA.mid(lineText, 1, sC - 1) & VBA.UCase$(VBA.mid(lineText, sC))
            ElseIf i = eL Then
                NewText = VBA.UCase$(VBA.mid(lineText, 1, eC - 1)) & VBA.mid(lineText, eC)
            Else
                NewText = VBA.UCase$(lineText)
            End If
            If NewText <> vbNullString Then Call Application.VBE.ActiveCodePane.codeModule.ReplaceLine(i, NewText)
        Next i
    End If
    Call Application.VBE.ActiveCodePane.SetSelection(sL, sC, eL, eC)
ErrorHandler:
    Select Case Err.Number
             Case 0
        Case Else
            Call WriteErrorLog("toUpperCase", False)
    End Select
End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : toLowerCase - convert selected code to lowercase
'* Created    : 18-02-2020 09:06
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub toLowerCase()
    On Error GoTo ErrorHandler
    Dim i           As Long
    Dim NewText     As String
    Dim lineText    As String
    Dim sL          As Long
    Dim eL          As Long
    Dim sC          As Long
    Dim eC          As Long

    Call Application.VBE.ActiveCodePane.GetSelection(sL, sC, eL, eC)

    If sL = eL Then
        lineText = Application.VBE.ActiveCodePane.codeModule.Lines(sL, 1)
        NewText = VBA.mid(lineText, 1, sC - 1) & VBA.LCase(VBA.mid(lineText, sC, eC - sC)) & VBA.mid(lineText, eC)
        If NewText <> vbNullString Then Call Application.VBE.ActiveCodePane.codeModule.ReplaceLine(sL, NewText)
    Else
        For i = sL To eL
            NewText = ""
            lineText = Application.VBE.ActiveCodePane.codeModule.Lines(i, 1)
            If i = sL Then
                NewText = VBA.mid(lineText, 1, sC - 1) & VBA.LCase(VBA.mid(lineText, sC))
            ElseIf i = eL Then
                NewText = VBA.LCase(VBA.mid(lineText, 1, eC - 1)) & VBA.mid(lineText, eC)
            Else
                NewText = VBA.LCase(lineText)
            End If
            If NewText <> vbNullString Then Call Application.VBE.ActiveCodePane.codeModule.ReplaceLine(i, NewText)
        Next i
    End If

    Call Application.VBE.ActiveCodePane.SetSelection(sL, sC, eL, eC)

ErrorHandler:
    Select Case Err.Number
             Case 0
        Case Else
            Call WriteErrorLog("toLowerCase", False)
    End Select
End Sub