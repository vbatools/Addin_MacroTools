VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBilderMsgBoxGenerator 
   Caption         =   "MsgBox Generator:"
   ClientHeight    =   8760.001
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8070
   OleObjectBlob   =   "frmBilderMsgBoxGenerator.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBilderMsgBoxGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : MsgBoxGenerator - MsgBox Builder
'* Created    : 15-09-2019 15:57
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Private Sub btnCancel_Click()
    Unload Me
End Sub
Private Sub UserForm_Activate()
    With Me
        .StartUpPosition = 0
        .Left = Application.Left + 0.5 * (Application.Width - .Width)
        .Top = Application.Top + 0.5 * (Application.Height - .Height)
    End With
    Const W         As Integer = 20
    Const H         As Integer = W
    With Application.CommandBars
        obtnCritical.Picture = .GetImageMso("CancelRequest", W, H)
        obtnQuestion.Picture = .GetImageMso("ButtonTaskSelfSupport", W, H)
        obtnCaution.Picture = .GetImageMso("LogicIncomplete", W, H)
        obtnInformation.Picture = .GetImageMso("Info", W, H)
    End With
    Me.lbHelp.Picture = Application.CommandBars.GetImageMso("Help", 18, 18)
End Sub
Private Sub txtMsg_Change()
    Call UnLocet(txtMsg, txtMsg2, lbClear2, lbStr2)
End Sub
Private Sub txtMsg2_Change()
    Call UnLocet(txtMsg2, txtMsg3, lbClear3, lbStr3)
End Sub
Private Sub txtMsg3_Change()
    Call UnLocet(txtMsg3, txtMsg4, lbClear4, lbStr4)
End Sub
Private Sub txtMsg4_Change()
    Call UnLocet(txtMsg4, txtMsg5, lbClear5, lbStr5)
End Sub
Private Sub UnLocet(ByRef txtMain As MSForms.textBox, ByRef txtChild As MSForms.textBox, ByRef lbClear As MSForms.Label, ByRef lbStr As MSForms.Label)
    With txtChild
        If txtMain.value = vbNullString Then
            .Enabled = False
            .value = vbNullString
            lbStr.Visible = True
            lbClear.Enabled = False
        Else
            .Enabled = True
            lbStr.Visible = False
            lbClear.Enabled = True
        End If
    End With
End Sub
Private Sub lbClearTitle_Click()
    txtTitel.value = vbNullString
End Sub
Private Sub lbClear1_Click()
    txtMsg.value = vbNullString
End Sub
Private Sub lbClear2_Click()
    txtMsg2.value = vbNullString
End Sub
Private Sub lbClear3_Click()
    txtMsg3.value = vbNullString
End Sub
Private Sub lbClear4_Click()
    txtMsg4.value = vbNullString
End Sub
Private Sub lbClear5_Click()
    txtMsg5.value = vbNullString
End Sub
'preview
Private Sub btnView_Click()
    Dim i           As Long
    Dim sSTR        As String
    sSTR = ButtonVal()
    i = CInt(Split(sSTR, "|")(0)) + CInt(Split(sSTR, "|")(1))
    If chbMsgBoxRtlReading Then i = i + vbMsgBoxRtlReading
    Call MsgBox(AddStringMsg(), i, txtTitel)
End Sub
Private Function ButtonVal() As String
    Dim iButton     As Integer
    Dim iButton1    As Integer

    iButton = vbOKOnly
    If obtnOKCancel Then iButton = vbOKCancel
    If obtnYesNo Then iButton = vbYesNo
    If obtnRepeatCancel Then iButton = vbRetryCancel
    If obtnYesNoCancel Then iButton = vbYesNoCancel
    If obtnObortRepeatIgnor Then iButton = vbAbortRetryIgnore

    iButton1 = 0
    If obtnCritical Then iButton1 = vbCritical
    If obtnQuestion Then iButton1 = vbQuestion
    If obtnCaution Then iButton1 = vbExclamation
    If obtnInformation Then iButton1 = vbInformation
    ButtonVal = iButton & "|" & iButton1
End Function
Private Function AddCodeText() As String
    Dim sSTR As String, sBtn As String, sVal As String, sTextMsg As String, sTitelMsg As String
    Dim sFerstSub As String, sEndSub As String
    Const sChr      As String = "||| & Chr(34) & |||"

    sVal = ButtonVal()

    Select Case CInt(Split(sVal, "|")(0))
        Case 0
            sBtn = "vbOKOnly"
            sFerstSub = "Call "
            sEndSub = vbNullString
        Case 1
            sBtn = "vbOKCancel"
            sFerstSub = "Call "
            sEndSub = vbNullString
        Case 4
            sBtn = "vbYesNo"
            sFerstSub = "If "
            sEndSub = " = vbNo Then" & " Exit Sub"
        Case 5
            sBtn = "vbRetryCancel"
            sFerstSub = "If "
            sEndSub = " = vbRetry Then" & vbNewLine & "End If"
        Case 3
            sBtn = "vbYesNoCancel"
            sFerstSub = "Select Case "
            sEndSub = vbNewLine & vbTab & "Case vbYes" & vbNewLine & vbTab & "Case vbNo" & vbNewLine & vbTab & "Case vbCancel" & vbNewLine & "End Select"
        Case 2
            sBtn = "vbAbortRetryIgnore"
            sFerstSub = "Select Case "
            sEndSub = vbNewLine & vbTab & "Case vbRetry" & vbNewLine & vbTab & "Case vbIgnore" & vbNewLine & vbTab & "Case vbAbort" & vbNewLine & "End Select"
    End Select

    Select Case CInt(Split(sVal, "|")(1))
        Case 0
            sBtn = sBtn
        Case 16
            sBtn = sBtn & "+vbCritical"
        Case 32
            sBtn = sBtn & "+vbQuestion"
        Case 48
            sBtn = sBtn & "+vbExclamation"
        Case 64
            sBtn = sBtn & "+vbInformation"
    End Select

    If chbMsgBoxRtlReading Then sBtn = sBtn & "+vbMsgBoxRtlReading"

    sTitelMsg = Replace(txtTitel.text, Chr(34), sChr)
    sTitelMsg = Replace(sTitelMsg, "|||", Chr(34))
    sTextMsg = AddStringMsg(True)

    sSTR = sFerstSub & "MsgBox(" & Chr(34)
    sSTR = sSTR & sTextMsg & Chr(34) & ", "
    sSTR = sSTR & sBtn & ", "
    sSTR = sSTR & Chr(34) & sTitelMsg & Chr(34) & ")" & sEndSub

    AddCodeText = sSTR
End Function

Private Sub lbInsertCode_Click()
    Dim sSTR As String, txtLine As String
    Dim iLine       As Integer

    sSTR = AddCodeText()
    If sSTR = vbNullString Then Exit Sub
    txtLine = SelectedLineColumnProcedure(Application.VBE.ActiveCodePane)
    If txtLine = vbNullString Then
        Me.Hide
        Exit Sub
    End If
    iLine = Split(txtLine, "|")(2)

    Call Application.VBE.ActiveCodePane.codeModule.InsertLines(iLine, sSTR)
    Me.Hide
End Sub
Private Function AddStringMsg(Optional bFlag As Boolean) As String
    Dim sTextMsg    As String
    Dim sSTR        As String
    Const sChr      As String = "||| & Chr(34) & |||"

    If bFlag Then
        sSTR = "||| & vbNewLine & |||"
    Else
        sSTR = vbNewLine
    End If

    sTextMsg = txtMsg.text
    If txtMsg2.value <> vbNullString Then sTextMsg = txtMsg.value & sSTR & txtMsg2.value
    If txtMsg3.value <> vbNullString Then sTextMsg = sTextMsg & sSTR & txtMsg3.value
    If txtMsg4.value <> vbNullString Then sTextMsg = sTextMsg & sSTR & txtMsg4.value
    If txtMsg5.value <> vbNullString Then sTextMsg = sTextMsg & sSTR & txtMsg5.value

    If bFlag Then
        sTextMsg = Replace(sTextMsg, Chr(34), sChr)
        sTextMsg = Replace(sTextMsg, "|||", Chr(34))
    End If

    AddStringMsg = sTextMsg
End Function