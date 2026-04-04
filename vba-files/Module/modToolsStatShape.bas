Attribute VB_Name = "modToolsStatShape"
Option Explicit
Option Private Module
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : I_StatisticShape - generate statistics of macros linked to shapes on an Excel sheet
'* Created    : 15-09-2019 15:48
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : AddShapeStatistic - create a table listing shapes, buttons in the file and their assigned macros
'* Created    : 22-03-2023 15:27
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub addShapeStatistic()

    Dim wb_name As String, macro_name As String
    Dim shp         As Shape
    Dim sh          As Worksheet
    Dim bDelSheet   As Boolean
    Dim i           As Integer
    Dim wb          As Workbook
    Const SH_SHAPE  As String = "SHAPES_VBA"

    If Not GetTargetWorkbook(wb, "Collecting shapes data:", "CREATE") Then Exit Sub
    On Error GoTo errMsg
    wb_name = wb.Name
    If wb_name = vbNullString Then Exit Sub

    Application.ScreenUpdating = False
    ActiveWorkbook.Sheets.Add After:=Sheets(Sheets.Count)
    Set wb = Workbooks(wb_name)

    On Error Resume Next
    Application.DisplayAlerts = False
    ActiveWorkbook.Worksheets(SH_SHAPE).Delete
    On Error GoTo 0

    With ActiveWorkbook.Worksheets.Add
        .Name = SH_SHAPE
        i = 2
        .Cells(i, 1).value = "Sheet Name"
        .Cells(i, 2).value = "Shape Name"
        .Cells(i, 3).value = "Shape Text"
        .Cells(i, 4).value = "Macro Name"
        For Each sh In wb.Worksheets
            For Each shp In sh.Shapes
                i = i + 1
                .Hyperlinks.Add Anchor:=Cells(i, 1), Address:="", SubAddress:=sh.Name & "!A1", TextToDisplay:=sh.Name
                .Cells(i, 2).value = shp.Name

                Select Case shp.Type
                        Case msoAutoShape
                        .Cells(i, 3).value = shp.TextFrame2.TextRange.Characters.text
                    Case msoFormControl, msoOLEControlObject
                        .Cells(i, 3).value = shp.AlternativeText
                    Case Else
                        .Cells(i, 3).value = "no"
                End Select
                On Error Resume Next
                macro_name = shp.OnAction
                On Error GoTo 0
                If macro_name = vbNullString Then
                    .Cells(i, 4).value = "no macro"
                Else
                    bDelSheet = True
                    .Cells(i, 4).value = Split(shp.OnAction, "!")(1)
                End If
            Next
        Next
        .Columns("A:D").EntireColumn.AutoFit
        .Cells(1, 1).value = wb.FullName
    End With
    If Not bDelSheet Then
        Application.DisplayAlerts = False
        ActiveWorkbook.Sheets(SH_SHAPE).Delete
        Application.DisplayAlerts = True
        Call MsgBox("No macro-related objects found", vbInformation)
    End If
    Application.ScreenUpdating = True
    Exit Sub
errMsg:
    If Err.Number = 1004 Then
        Application.DisplayAlerts = False
        ActiveWorkbook.Sheets(SH_SHAPE).Delete
        Application.DisplayAlerts = True
        ActiveSheet.Name = SH_SHAPE
        Err.Clear
        Resume Next
    Else
        Application.ScreenUpdating = True
        Call WriteErrorLog("addShapeStatistic", True)
    End If
End Sub