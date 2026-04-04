VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBilderFormat 
   Caption         =   "Format Builder:"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10455
   OleObjectBlob   =   "frmBilderFormat.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBilderFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'***********************************************************************************************************
'* Author     : frmBilderFormat - Format Builder for strings
'* Created    : 15.09.2019
'* Author     : VBATools
'* Copyright  : Apache License
'***********************************************************************************************************

Private arrFormatDate As Variant
Private arrFormatDateDiscription As Variant
Private arrFormatValue As Variant
Private arrFormatValueDiscription As Variant
Private arrFormatDateCustom As Variant
Private arrFormatDateCustomDiscription As Variant

Private Function AddCode() As String
    Dim sSTR As String, sErr As String

    If obtnFormat Then
        sSTR = cmbFormat.value
    Else
        sSTR = LTrim(txtCustom.text)
    End If

    If lbView.Caption = vbNullString Then sErr = "Input field is empty!" & vbNewLine
    If lbView.Caption Like "Error: *" Then sErr = sErr & "Error in source format"
    If sErr <> vbNullString Then
        Call MsgBox(sErr, vbCritical, "Error:")
        Exit Function
    End If

    sSTR = "VBA.Format$(" & Replace(txtValue.value, ",", ".") & ", " & Chr(34) & sSTR & Chr(34) & ")"
    AddCode = sSTR
End Function

Private Sub lbInsertCode_Click()
    Dim iLine       As Integer
    Dim txtCode As String, txtLine As String
    'get code
    txtCode = AddCode()
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
Private Sub cmbCancel_Click()
    Me.Hide
End Sub

Private Sub cmbCustomFormat_Change()
    If cmbCustomFormat.ListIndex = -1 Then Exit Sub
    If obtnCustomFormat Then
        txtDiscription.text = arrFormatDateCustomDiscription(cmbCustomFormat.ListIndex)
    End If
End Sub
Private Sub txtCustom_Change()
    Call AddFormat(txtCustom.text)
End Sub
Private Sub lbClear_Click()
    txtCustom.text = vbNullString
End Sub
Private Sub cmbFormat_Change()
    If cmbFormat.ListIndex = -1 Then Exit Sub
    If obtnDate And obtnFormat Then
        txtDiscription.text = arrFormatDateDiscription(cmbFormat.ListIndex)
    Else
        txtDiscription.text = arrFormatValueDiscription(cmbFormat.ListIndex)
    End If

    Call AddFormat(cmbFormat.value)
End Sub

Private Sub lbAddCustom_Click()
    If cmbCustomFormat.value <> vbNullString Then
        txtCustom.text = txtCustom & " " & cmbCustomFormat.value
    End If
End Sub

Private Sub obtnDate_Change()
    Call AddList
End Sub

Private Sub AddList()
    If obtnFormat Then
        cmbCustomFormat.Clear
        If obtnDate Then
            cmbFormat.List = arrFormatDate
        Else
            cmbFormat.List = arrFormatValue
        End If
        obtnValue.Visible = True
    Else
        cmbFormat.Clear
        cmbCustomFormat.List = arrFormatDateCustom
        obtnValue.Visible = False
    End If
    Call ChengeFlag(obtnFormat)
End Sub
Private Sub ChengeFlag(ByVal Flag As Boolean)
    obtnDate.Visible = Flag
    cmbFormat.Visible = Flag
    cmbCustomFormat.Visible = (Not Flag)
    txtCustom.Visible = (Not Flag)
    lbClear.Visible = (Not Flag)
    lbAddCustom.Visible = (Not Flag)
    If Flag Then
        Frame2.Left = lbAddCustom.Left - 5
    Else
        Frame2.Left = cmbCustomFormat.Left + cmbCustomFormat.Width + 5
    End If
End Sub
Private Sub obtnFormat_Click()
    Call AddList
End Sub
Private Sub obtnCustomFormat_Click()
    Call AddList
End Sub
Private Sub txtValue_Change()
    If cmbFormat.value = vbNullString Then
        Call AddFormat(cmbCustomFormat.value)
    Else
        Call AddFormat(cmbFormat.value)
    End If
End Sub
Private Sub AddFormat(ByVal sSTR As String)
    On Error GoTo err_msg
    lbView.Caption = Format(CDbl(txtValue.value), sSTR)
    lbView.ForeColor = &H8000000D
    Exit Sub
err_msg:
    Select Case Err.Number
             Case 6
            lbView.Caption = "Error: enter a smaller number!"
        Case 13
            lbView.Caption = vbNullString
        Case Else
            lbView.Caption = "Error:" & Err.Description & " " & Err.Number
    End Select
    lbView.ForeColor = &H8080FF
    Err.Clear
End Sub

Private Sub txtValue_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim txt         As String
    txt = Me.txtValue    ' read text from field (to prevent entering two or more commas)
    Select Case KeyAscii
             Case 8:    ' Backspace pressed - do nothing
        Case 44: KeyAscii = IIf(InStr(1, txt, ",") > 0, 0, 44)    ' if comma already exists - cancel character input
        Case 46: KeyAscii = IIf(InStr(1, txt, ",") > 0, 0, 44)    ' replace dot with comma on input
        Case 48 To 57    ' if a digit is entered - do nothing
        Case Else: KeyAscii = 0    ' otherwise cancel character input
    End Select
End Sub

Private Sub UserForm_Initialize()
    With Me
        .StartUpPosition = 0
        .Left = Application.Left + 0.5 * (Application.Width - .Width)
        .Top = Application.Top + 0.5 * (Application.Height - .Height)
    End With
    arrFormatDate = Array("General Date", "Long Date", "Medium Date", "Short Date", "Long Time", "Medium Time", "Short Time")
    arrFormatDateDiscription = Array("Display date and/or time, e.g., 4/3/93 05:34 PM. If there is no fractional part, only the date is displayed, e.g., 4/3/93. If there is no integer part, only the time is displayed, e.g., 05:34 PM. Date display is determined by system settings.", _
            "Display date using the system's long date format.", _
            "Display date using the medium date format appropriate for the language version of the host application.", _
            "Display date using the system's short date format.", _
            "Display time using the system's long time format; includes hours, minutes, and seconds.", _
            "Display time in 12-hour format using hours, minutes, and AM/PM indicator.", _
            "Display time in 24-hour format, e.g., 17:45.")
    arrFormatValue = Array("General Number", "Currency", "Fixed", "Standard", "Percent", "Scientific", "Yes/No", "True/False", "On/Off")
    arrFormatValueDiscription = Array("Display number without thousands separator.", _
            "Display number with thousands separator if needed; displays two digits to the right of the decimal separator. Output is based on system locale settings.", _
            "Display at least one digit to the left and two digits to the right of the decimal separator.", _
            "Display number with thousands separator; displays at least one digit to the left and two digits to the right of the decimal separator.", _
            "Display number multiplied by 100 with a percent sign (%) appended to the right; always displays two digits to the right of the decimal separator.", _
            "Use standard exponential notation.", _
            "Display No if number equals 0; otherwise display Yes.", _
            "Display False if number equals 0; otherwise display True.", _
            "Display Off if number equals 0; otherwise display On.")
    arrFormatDateCustom = Array("c", "d", "dd", "ddd", "dddd", "ddddd", "dddddd", "w", "ww", "m", "mm", "mmm", "mmmm", "q", "y", "yy", "yyyy", "h", "hh", "n", "nn", "s", "ss", "ttttt")
    arrFormatDateCustomDiscription = Array("Date separator. Some locales may use different characters to represent the date separator. This separator separates day, month, and year when date values are formatted. The character used as the date separator in formatted output is determined by system settings.", _
            "Display day as a number without a leading zero (1–31).", _
            "Display day as a number with a leading zero (01–31).", _
            "Display day using abbreviations (Sun–Sat). Localized.", _
            "Display day using full name (Sunday–Saturday). Localized.", _
            "Display date using full format (including day, month, and year) corresponding to the system's short date format. The default short date format is m/d/yy.", _
            "Display date number using full format (including day, month, and year) corresponding to the system's long date format. The default long date format is mmmm dd, yyyy.", _
            "Display day of week as a number (1 for Sunday through 7 for Saturday).", _
            "Display week of year as a number (1–54).", _
            "Display month as a number without a leading zero (1–12). If m immediately follows h or hh, the minute rather than the month is displayed.", _
            "Display month as a number with a leading zero (01–12). If m immediately follows h or hh, the minute rather than the month is displayed.", _
            "Display month using abbreviations (Jan–Dec). Localized.", _
            "Display month using full name (January–December). Localized.", _
            "Display quarter of year as a number (1–4).", "Display day of year as a number (1–366).", "Display year as a 2-digit number (00–99).", _
            "Display year as a 4-digit number (100–9999).", "Display hour as a number without a leading zero (0–23).", _
            "Display hour as a number with a leading zero (00–23).", "Display minute as a number without a leading zero (0–59).", _
            "Display minute as a number with a leading zero (00–59).", _
            "Display second as a number without a leading zero (0–59).", _
            "Display second as a number with a leading zero (00–59).", _
            "Display time in full format (including hour, minute, and second) using the time separator defined in the system's time format. A leading zero is displayed if the leading zero option is selected and the time is before 10:00 A.M. or P.M. The default time format is h:mm:ss.")
    cmbFormat.List = arrFormatDate
    Call ChengeFlag(obtnFormat)
    Me.lbHelp.Picture = Application.CommandBars.GetImageMso("Help", 18, 18)
End Sub