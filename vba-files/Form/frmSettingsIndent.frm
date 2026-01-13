VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSettingsIndent 
   Caption         =   "Настройки:"
   ClientHeight    =   8550.001
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13950
   OleObjectBlob   =   "frmSettingsIndent.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSettingsIndent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Option Compare Text
Option Base 1
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : OptionsCodeFormat - настройка форматирования кода, растановка отступов
'* Created    : 15-09-2019 15:57
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Private Sub cmbCancel_Click()
    ThisWorkbook.Save
    Unload Me
End Sub
Private Sub lbCancel_Click()
    Call cmbCancel_Click
End Sub
Private Sub UserForm_Activate()
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.top = Application.top + (0.5 * Application.Height) - (0.5 * Me.Height)

    Dim OptionsTb   As ListObject
    Set OptionsTb = shSettings.ListObjects(modAddinConst.TB_OPTIONS_IDEDENT)
    Call UpdateCodeListBox
    With OptionsTb.ListColumns(2)
        SpinBtnTab.Value = .Range(2, 1)
        txtTab.Value = .Range(2, 1)
        chbIndentProc.Value = .Range(3, 1)
        chbIndentFirst.Value = .Range(4, 1)
        chbIndentDim.Value = .Range(5, 1)
        chbIndentCmt.Value = .Range(6, 1)
        chbIndentCase.Value = .Range(7, 1)
        chbAlignCont.Value = .Range(8, 1)
        chbAlignIgnoreOps.Value = .Range(9, 1)
        chbDebugCol1.Value = .Range(10, 1)
        chbAlignDim.Value = .Range(11, 1)
        SpinBtn.Value = .Range(12, 1)
        txtmiAlignDimCol.Value = .Range(12, 1)
        chbCompilerStuffCol1.Value = .Range(13, 1)
        chbIndentCompilerStuff.Value = .Range(14, 1)
        SpinBtnComment.Value = .Range(16, 1)
        txtComment.Value = .Range(16, 1)

        Select Case .Range(15, 1)
            Case "Absolute":
                obtnAbsolute.Value = True
            Case "SameGap":
                obtnSameGap.Value = True
            Case "StandardGap":
                obtnStandardGap.Value = True
            Case "AlignInCol":
                obtnAlignInCol.Value = True
        End Select
    End With
    Me.lbHelp.Picture = Application.CommandBars.GetImageMso("Help", 18, 18)
End Sub
Private Sub txtTab_Change()
    Call SetOptFromTable(2, txtTab.Value)
End Sub
Private Sub chbIndentProc_Change()
    chbIndentFirst.Enabled = chbIndentProc.Value
    chbIndentDim.Enabled = chbIndentProc.Value
    Call SetOptFromTable(3, chbIndentProc.Value)
End Sub
Private Sub chbIndentFirst_Change()
    Call SetOptFromTable(4, chbIndentFirst.Value)
End Sub
Private Sub chbIndentDim_Change()
    Call SetOptFromTable(5, chbIndentDim.Value)
End Sub
Private Sub chbIndentCmt_Change()
    Call SetOptFromTable(6, chbIndentCmt.Value)
End Sub
Private Sub chbIndentCase_Change()
    Call SetOptFromTable(7, chbIndentCase.Value)
End Sub
Private Sub chbAlignCont_Change()
    chbAlignIgnoreOps.Value = Not chbAlignCont.Value
    Call SetOptFromTable(8, chbAlignCont.Value)
End Sub
Private Sub chbAlignIgnoreOps_Change()
    If chbAlignIgnoreOps Then chbAlignCont.Value = False
    Call SetOptFromTable(9, chbAlignIgnoreOps.Value)
End Sub
Private Sub chbDebugCol1_Change()
    Call SetOptFromTable(10, chbDebugCol1.Value)
End Sub
Private Sub chbAlignDim_Change()
    txtmiAlignDimCol.Enabled = chbAlignDim.Value
    SpinBtn.Enabled = chbAlignDim.Value
    Call SetOptFromTable(11, chbAlignDim.Value)
End Sub
Private Sub txtmiAlignDimCol_Change()
    Call SetOptFromTable(12, txtmiAlignDimCol.Value)
End Sub
Private Sub SpinBtn_SpinDown()
    Call SpinBtnChange(0, 30, Me.SpinBtn, Me.txtmiAlignDimCol)
End Sub
Private Sub SpinBtn_SpinUp()
    Call SpinBtnChange(0, 30, Me.SpinBtn, Me.txtmiAlignDimCol)
End Sub
Private Sub SpinBtnTab_SpinDown()
    Call SpinBtnChange(4, 8, Me.SpinBtnTab, Me.txtTab)
End Sub
Private Sub SpinBtnTab_SpinUp()
    Call SpinBtnChange(4, 8, Me.SpinBtnTab, Me.txtTab)
End Sub
Private Sub SpinBtnComment_SpinDown()
    Call SpinBtnChange(0, 100, Me.SpinBtnComment, Me.txtComment)
End Sub
Private Sub SpinBtnComment_SpinUp()
    Call SpinBtnChange(0, 100, Me.SpinBtnComment, Me.txtComment)
End Sub
Private Sub SpinBtnChange(ByVal iMin As Byte, ByVal iMax As Byte, ByRef objSpinBtn As MSForms.SpinButton, ByRef objTxt As MSForms.textBox)
    With objSpinBtn
        If .Value < iMin Then .Value = iMin
        If .Value > iMax Then .Value = iMax
        objTxt.Text = .Value
    End With
End Sub
Private Sub chbCompilerStuffCol1_Change()
    chbIndentCompilerStuff.Value = Not chbCompilerStuffCol1.Value
    Call SetOptFromTable(13, chbCompilerStuffCol1.Value)
End Sub
Private Sub chbIndentCompilerStuff_Change()
    If chbIndentCompilerStuff Then chbCompilerStuffCol1.Value = False
    Call SetOptFromTable(14, chbIndentCompilerStuff.Value)
End Sub
Private Sub obtnAbsolute_Change()
    Call SetOptFromTable(15, obtnAbsolute.Tag)
End Sub
Private Sub obtnAlignInCol_Change()
    txtComment.Enabled = obtnAlignInCol.Value
    SpinBtnComment.Enabled = obtnAlignInCol.Value
    Call SetOptFromTable(15, obtnAlignInCol.Tag)
End Sub
Private Sub obtnSameGap_Change()
    Call SetOptFromTable(15, obtnSameGap.Tag)
End Sub
Private Sub obtnStandardGap_Change()
    Call SetOptFromTable(15, obtnStandardGap.Tag)
End Sub
Private Sub txtComment_Change()
    Call SetOptFromTable(16, txtComment.Value)
End Sub
Private Sub UpdateCodeListBox()

    Dim asCodeLines(1 To 30) As String
    Dim i           As Integer

    'Define the example procedure code lines
    asCodeLines(1) = "' Example Procedure"
    asCodeLines(2) = "Sub ExampleProc()"
    asCodeLines(3) = ""
    asCodeLines(4) = "'надстройка " & modAddinConst.NAME_ADDIN
    asCodeLines(5) = "'© 2018-" & VBA.Year(Now()) & " by " & modAddinConst.NAME_ADDIN & " Ltd."
    asCodeLines(6) = ""
    asCodeLines(7) = "Dim iCount As Integer"
    asCodeLines(8) = "Static sName As String"
    asCodeLines(9) = ""
    asCodeLines(10) = "If YouWantMoreExamplesAndTools Then"
    asCodeLines(11) = "' Visit https://VBATools.ru"
    asCodeLines(12) = "' https://vk.com/vbatools"
    asCodeLines(13) = "Select Case X"
    asCodeLines(14) = "Case ""A"""
    asCodeLines(15) = "' If you have any comments or suggestions, _"
    asCodeLines(16) = " or find valid VBA code that isn't indented correctly,"
    asCodeLines(17) = ""
    asCodeLines(18) = "#If VBA6 Then"
    asCodeLines(19) = "MsgBox ""Contact info@VBATools.ru"""
    asCodeLines(20) = "#End If"
    asCodeLines(21) = ""
    asCodeLines(22) = "Case ""Continued strings and parameters can be"" _"
    asCodeLines(23) = "& ""lined up for easier reading, optionally ignoring"" _"
    asCodeLines(24) = ", ""any operators (&+, etc) at the start of the line."""
    asCodeLines(25) = ""
    asCodeLines(26) = "Debug.Print ""X<>1"""
    asCodeLines(27) = "End Select           'Case X"
    asCodeLines(28) = "End If               'More Tools?"
    asCodeLines(29) = ""
    asCodeLines(30) = "End Sub"


    'Run the array through the indenting code
    RebuildCodeArray asCodeLines, "", 0

    'Put the procedure code in the list box.

    txtCode.Text = vbNullString
    For i = LBound(asCodeLines) To UBound(asCodeLines)
        If i = UBound(asCodeLines) Then
            txtCode.Text = txtCode.Text & asCodeLines(i)
        Else
            txtCode.Text = txtCode.Text & asCodeLines(i) & vbNewLine
        End If
    Next
End Sub

Private Sub SetOptFromTable(ByVal iRow As Byte, ByVal iVal As Variant)
    Dim OptionsTb   As ListObject
    Set OptionsTb = shSettings.ListObjects(modAddinConst.TB_OPTIONS_IDEDENT)
    OptionsTb.ListColumns(2).Range(iRow, 1) = iVal
    Call UpdateCodeListBox
End Sub
