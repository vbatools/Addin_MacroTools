VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBilderProcedure 
   Caption         =   "Prosedure Bilder:"
   ClientHeight    =   8265.001
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   17805
   OleObjectBlob   =   "frmBilderProcedure.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBilderProcedure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : BilderProcedure - Procedure Builder
'* Created    : 15-09-2019 15:57
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Private Sub btnCancel_Click()
    Me.Hide
End Sub

Private Sub lbCancel_Click()
    Call btnCancel_Click
End Sub

Private Sub UserForm_Initialize()
    With Me
        .StartUpPosition = 0
        .Left = Application.Left + 0.5 * (Application.Width - .Width)
        .Top = Application.Top + 0.5 * (Application.Height - .Height)
    End With
    With cmbFunc
        .AddItem "Boolean"
        .AddItem "String"
        .AddItem "Byte"
        .AddItem "Integer"
        .AddItem "Long"
        .AddItem "Single"
        .AddItem "Double"
        .AddItem "Currency"
        .AddItem "Variant"
        .AddItem "Date"
        .AddItem "Object"
    End With
    txtErroName.text = "<- Input field" & Chr(34) & Replace(lbName.Caption, "*:", vbNullString) & Chr(34) & "must be filled!"
End Sub
Private Sub UserForm_Activate()
    Me.lbHelp.Picture = Application.CommandBars.GetImageMso("Help", 18, 18)
End Sub
Private Sub chbAll_Change()
    Dim Flag        As Boolean
    Flag = chbAll.value
    chbScreen.value = Flag
    chbCalculations.value = Flag
    chbAlerts.value = Flag
    chbEvents.value = Flag
    chbMsg.value = Flag
    chbUseDefaultMsg.value = Flag
End Sub
Private Sub optTypeModif_Change()
    txtViewCode.text = AddCode
End Sub
Private Sub txtName_Change()
    Dim txt         As String
    If txtName = vbNullString Then
        txtName.BorderColor = &HC0C0FF
    Else
        txtName.BorderColor = &H8000000D
    End If
    txtViewCode.text = AddCode
    txt = txtName.text
    If VBA.Left$(txt, 1) = "_" Then
        txt = VBA.Right(txt, VBA.Len(txt) - 1)
        txtName.text = txt
    End If
End Sub
Private Sub txtName_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim sTemplate As String, txt As String
    txt = txtName.text
    sTemplate = "!@#$%^&*+=.,'№/\|-:;{}[]() <>" & Chr(34)
    If InStr(1, sTemplate, ChrW(KeyAscii)) > 0 Then KeyAscii = 0
    If txt = vbNullString Then
        Select Case KeyAscii
                 Case 48 To 57: KeyAscii = 0
        End Select
    End If
    If VBA.Left$(txt, 1) = "_" Then
        txt = VBA.Right(txt, VBA.Len(txt) - 1)
        txtName.text = txt
    End If
End Sub
Private Sub cmbFunc_Change()
    Call AddBackColorCombobox
    txtViewCode.text = AddCode
End Sub
Private Sub optTypeProcedure_Change()
    cmbFunc.Enabled = Not optTypeProcedure.value
    chbArray.Enabled = Not optTypeProcedure.value
    Call AddBackColorCombobox
    txtViewCode.text = AddCode
End Sub
Private Sub AddBackColorCombobox()
    cmbFunc.BorderColor = &H8000000D
    If (Not optTypeProcedure) Then
        If cmbFunc.value = vbNullString Then
            cmbFunc.BorderColor = &HC0C0FF
        End If
    End If
End Sub

Private Sub TernOffOn()
    If (chbScreen.value + chbCalculations.value + chbAlerts.value + chbEvents.value) <> 0 Then
        chbAddMainProceure.Enabled = True
    Else
        chbAddMainProceure.Enabled = False
    End If
End Sub

Private Sub chbArray_Change()
    txtViewCode.text = AddCode
End Sub
Private Sub chbAlerts_Change()
    txtViewCode.text = AddCode
    Call TernOffOn
End Sub
Private Sub chbCalculations_Change()
    txtViewCode.text = AddCode
    Call TernOffOn
End Sub
Private Sub chbEvents_Change()
    txtViewCode.text = AddCode
    Call TernOffOn
End Sub
Private Sub chbMsg_Change()
    txtViewCode.text = AddCode
    chbUseDefaultMsg.Enabled = chbMsg.value
    txtMsg.Enabled = chbMsg.value
End Sub
Private Sub chbScreen_Change()
    txtViewCode.text = AddCode
    Call TernOffOn
End Sub
Private Sub txtMsg_Change()
    txtViewCode.text = AddCode
End Sub
Private Sub txtDiscprition_Change()
    txtViewCode.text = AddCode
End Sub
Private Sub optDefaultError_Change()
    txtViewCode.text = AddCode
End Sub
Private Sub optResumNext_Change()
    txtViewCode.text = AddCode
End Sub
Private Sub chbUseDefaultMsg_Change()
    txtViewCode.text = AddCode
End Sub
Private Sub chbOffDiscription_Change()
    txtViewCode.text = AddCode
    txtDiscprition.Enabled = chbOffDiscription.value
End Sub
Private Function AddCode() As String
    Dim strCode As String, strSpes As String, strEndLine As String
    Dim TypeModif As String, TypeProc As String, strDiscprition As String
    Dim TypeFunction As String, ResultDimFunc As String, ResultEndFunc As String
    Dim strMsg As String, strMsg1 As String, CustMsg As String
    Dim ErrorMsgFerst As String, ErrorMsgEnd As String
    Dim MsgStop     As String
    Dim ScreenUpdatingCalculationTrue As String, ScreenUpdatingCalculationFalse As String
    Dim txtArray    As String

    If txtName.text = vbNullString Then
        txtErroName.Visible = True
        Exit Function
    Else
        txtErroName.Visible = False
    End If
    If (Not optTypeProcedure) Then
        If cmbFunc.value = vbNullString Then
            MsgStop = "Function data type selection field must be filled!"
        End If
    End If

    If MsgStop <> vbNullString Then
        Call MsgBox(MsgStop, vbOKOnly + vbCritical, "Error:")
        Exit Function
    End If

    strEndLine = vbNewLine & vbTab
    strSpes = Space(1)
    ScreenUpdatingCalculationTrue = "Call ScreenUpdatingCalculation(Screen:=True, Calculat:=True, Alerts:=True, Events:=True)"

    'access modifier type
    If optTypeModif Then
        TypeModif = "Public"
    Else
        TypeModif = "Private"
    End If

    'array for functions
    If (Not optTypeProcedure.value) And chbArray.value Then
        txtArray = " ()"
    End If

    'procedure or function
    If optTypeProcedure Then
        TypeProc = "Sub"
        TypeFunction = vbNullString
    Else
        TypeProc = "Function"
        TypeFunction = " as " & cmbFunc.value
        ResultDimFunc = vbNewLine & vbTab & "Dim Result" & txtArray & " as " & cmbFunc.value
        ResultEndFunc = vbNewLine & vbTab & txtName.text & " = Result"
    End If

    'disable description
    If chbOffDiscription Then
        strDiscprition = "'" & addStringDelimetr() & vbNewLine
        strDiscprition = strDiscprition & "'" & TypeProcedyreComments(TypeProc) & vbTab & txtName.text & " - " & txtDiscprition.value & vbNewLine
        strDiscprition = strDiscprition & addArrFromTBComments()(0, 2) & vbNewLine
        strDiscprition = strDiscprition & "'" & addStringDelimetr() & vbNewLine
    End If

    'display message upon completion
    If chbMsg Then
        Dim txtNewLine As String
        If txtMsg.text <> vbNullString Then txtNewLine = " & vbNewLine & "
        If chbUseDefaultMsg Then strMsg1 = Chr(34) & "Execution " & txtName.text & " completed!" & Chr(34) & txtNewLine
        CustMsg = txtMsg.text
        If CustMsg = vbNullString Then
            If chbUseDefaultMsg Then
                CustMsg = vbNullString
            Else
                CustMsg = Chr(34) & vbNullString & Chr(34)
            End If
        Else
            CustMsg = Replace(CustMsg, Chr(34), "| & Chr(34) & |")
            CustMsg = Chr(34) & Replace(CustMsg, "|", Chr(34)) & Chr(34)
        End If
        strMsg1 = strMsg1 & CustMsg
        strMsg = strEndLine & "Call MsgBox(" & strMsg1 & ", vbOKOnly + vbInformation," & Chr(34) & txtName.text & Chr(34) & ")"
    End If

    'error handling
    If optDefaultError Then
        ErrorMsgFerst = vbNullString
        ErrorMsgEnd = vbNullString
    End If

    If optResumNext Then
        ErrorMsgFerst = strEndLine & "On Error Resume Next"
        ErrorMsgEnd = vbNullString
    End If
    ScreenUpdatingCalculationFalse = "Call ScreenUpdatingCalculation(Screen:=" & (Not chbScreen.value) & ", Calculat:=" & (Not chbCalculations.value) & ", Alerts:=" & (Not chbAlerts.value) & ", Events:=" & (Not chbEvents.value) & ")"

    'if nothing is disabled, then nothing to enable
    If ScreenUpdatingCalculationFalse = ScreenUpdatingCalculationTrue Then
        ScreenUpdatingCalculationTrue = vbNullString
        ScreenUpdatingCalculationFalse = vbNullString
    End If

    If optErrorHandele Then
        ErrorMsgFerst = strEndLine & "On Error GoTo ErrorHandler"
        ErrorMsgEnd = strEndLine & "Exit " & TypeProc & vbNewLine & "ErrorHandler:" & strEndLine & ScreenUpdatingCalculationTrue
        ErrorMsgEnd = ErrorMsgEnd & strEndLine & "Select Case Err"
             ErrorMsgEnd = ErrorMsgEnd & strEndLine & vbTab & Chr(39) & "error handling for use uncomment"
        ErrorMsgEnd = ErrorMsgEnd & strEndLine & vbTab & Chr(39) & "Case"
        ErrorMsgEnd = ErrorMsgEnd & strEndLine & vbTab & "Case Else:"
        ErrorMsgEnd = ErrorMsgEnd & strEndLine & vbTab & vbTab & "Debug.Print " & Chr(34) & "An error occurred in" & txtName & Chr(34) & " & vbNewLine & Err.Number & vbNewLine & Err.Description"
        ErrorMsgEnd = ErrorMsgEnd & strEndLine & "End Select"
    End If

    'code generation
    strCode = strDiscprition
    strCode = strCode & TypeModif & strSpes & TypeProc & strSpes & txtName.value & strSpes & "()" & TypeFunction & txtArray
    strCode = strCode & ResultDimFunc
    strCode = strCode & ErrorMsgFerst
    strCode = strCode & strEndLine & ScreenUpdatingCalculationFalse
    strCode = strCode & strEndLine & strEndLine & Chr(39) & "placeholder for code" & strEndLine
    strCode = strCode & ResultEndFunc
    strCode = strCode & strEndLine & ScreenUpdatingCalculationTrue
    strCode = strCode & strMsg
    strCode = strCode & ErrorMsgEnd
    strCode = strCode & vbNewLine & "End " & TypeProc

    AddCode = strCode
End Function

Private Function AddMainProceure() As String
    Dim txtCode     As String
    txtCode = txtViewCode.text

    'copy ScreenUpdatingCalculation procedure
    If chbAddMainProceure Then
        Dim snippets As ListObject
        Dim i_row   As Long
        Set snippets = shSettings.ListObjects(TB_SNIPETS)
        i_row = snippets.ListColumns(2).DataBodyRange.Find(What:="ScreenUpdatingCalculation", LookIn:=xlValues, LookAt:=xlWhole).Row
        txtCode = txtCode & vbNewLine & snippets.Range(i_row, 3)
    End If
    AddMainProceure = txtCode
End Function

Private Sub lbInsertCode_Click()
    Dim iLine       As Integer
    Dim txtCode As String, txtLine As String
    'get code
    txtCode = AddMainProceure()
    If txtCode = vbNullString Then Exit Sub
    txtLine = SelectedLineColumnProcedure(Application.VBE.ActiveCodePane)
    If txtLine = vbNullString Then
        Me.Hide
        Exit Sub
    End If
    iLine = Split(txtLine, "|")(2)

    Call Application.VBE.ActiveCodePane.codeModule.InsertLines(iLine, txtCode)
    Me.Hide
End Sub