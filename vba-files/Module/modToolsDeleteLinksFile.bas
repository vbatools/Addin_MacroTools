Attribute VB_Name = "modToolsDeleteLinksFile"
Option Explicit
Option Private Module

' Global variable to store the report workbook
Dim g_ResultBook    As Workbook

' Constants for report formatting
Private Const COLOR_HEADER_BG As Long = 12611584
Private Const COLOR_HEADER_TEXT As Long = 16777215
Private Const COLOR_SUBHEADER_BG As Long = 16247773

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : ExternalLinkUtility - Entry point. Main procedure for searching links.
'* Created    : 16-06-2023 15:07
'* Author     : VBATools (Refactored)
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub ExternalLinkUtility()
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Call ReportExternalLinks(ActiveWorkbook)

CleanUp:
    Application.ScreenUpdating = True
    If Not g_ResultBook Is Nothing Then
        g_ResultBook.Activate
        Set g_ResultBook = Nothing
    End If
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume CleanUp
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : OutputLinkInfo - Output information to report
'* Created    : 16-06-2023 15:10
'* Argument(s): typ (Type), wbk (Workbook path), wsh (Sheet), loc (Location),
'*              adr (Address), fml (Formula), txt (Note)
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function OutputLinkInfo(ByVal typ As String, ByVal wbk As String, ByVal wsh As String, _
        ByVal loc As String, ByVal adr As String, ByVal fml As String, _
        ByVal txt As String)
    Static resultLn As Long

    ' Initialize report workbook on first call
    If g_ResultBook Is Nothing Then
        Set g_ResultBook = Workbooks.Add
        With g_ResultBook.Worksheets.item(1)
            ' Report header
            With .Range("A1:F1")
                .value = "External Links Report"
                .Font.Bold = True
                .Font.Size = 18
                .Interior.Color = COLOR_HEADER_BG
                .Font.Color = COLOR_HEADER_TEXT
                .Merge
            End With

            ' Column headers
            .Range("A2").value = "Type"
            .Range("B2").value = "Workbook"
            .Range("C2").value = "Sheet"
            .Range("D2").value = "Location"
            .Range("E2").value = "Link/Formula"
            .Range("F2").value = "Fix Instructions"

            With .Range("A2:F2")
                .Interior.Color = COLOR_SUBHEADER_BG
                .Font.Bold = True
            End With

            ' Configure column widths
            .Columns("A").ColumnWidth = 22
            .Columns("B").ColumnWidth = 15
            .Columns("C").ColumnWidth = 28
            .Columns("D").ColumnWidth = 28
            .Columns("E").ColumnWidth = 60
            .Columns("F").ColumnWidth = 60

            ' Add filter
            .Range("A2:F2").AutoFilter
        End With
        resultLn = 2
    End If

    ' Write data
    resultLn = resultLn + 1
    With g_ResultBook.Worksheets.item(1)
        .Range("A" & resultLn).value = typ
        .Range("B" & resultLn).value = Dir(wbk)
        .Range("C" & resultLn).value = wsh
        .Range("D" & resultLn).value = loc

        ' Add hyperlink, if possible
        If (Len(adr) > 0) And (Len(Dir(wbk)) > 0) Then
            .Hyperlinks.Add .Range("D" & resultLn), wbk, "'" & wsh & "'!" & adr, "Go to Issue", loc
        End If

        ' Add apostrophe to display formula as text
        .Range("E" & resultLn).value = "'" & fml
        .Range("F" & resultLn).value = txt
    End With
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub         : ReportExternalLinks - Main logic for searching links in workbook
'* Argument(s): wkbk - Workbook to check
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub ReportExternalLinks(wkbk As Excel.Workbook)
    Dim wksht       As Excel.Worksheet
    Dim numLinks    As Long
    numLinks = 0

    ' ==========================================
    ' WORKSHEET-LEVEL CHECK
    ' ==========================================
    For Each wksht In wkbk.Worksheets

        ' 1. Search for links in cell formulas
        ' Optimization: Use Find instead of iterating through all cells
        Call CheckCellFormulas(wksht, wkbk, numLinks)

        ' 2. Search for links in shapes
        Call CheckShapeLinks(wksht, wkbk, numLinks)

        ' 3. Search for links in conditional formatting
        Call CheckConditionalFormatting(wksht, wkbk, numLinks)

        ' 4. Search for links in charts
        Call CheckChartLinks(wksht, wkbk, numLinks)

        ' 5. Search for links in pivot tables
        Call CheckPivotTableLinks(wksht, wkbk, numLinks)

        ' 6. Search for links in data validation
        Call CheckDataValidationLinks(wksht, wkbk, numLinks)

    Next wksht

    ' ==========================================
    ' WORKBOOK-LEVEL CHECK
    ' ==========================================

    ' 7. Search and clean up links in named ranges
    Call CheckNamedRangeLinks(wkbk, numLinks)

    ' ==========================================
    ' COMPLETION AND REPORT
    ' ==========================================
    Application.ScreenUpdating = True

    If numLinks <= 0 Then
        MsgBox "Check completed:" & vbCrLf & vbCrLf & _
                "No external links detected: " & Dir(wkbk.FullName), vbInformation
    Else
        Dim msg     As String
        msg = "Check completed:" & vbCrLf & vbCrLf
        ' If broken links were deleted, numLinks includes the final deletion record,
        ' so we display numLinks - 1 for found links if delCt > 0
        ' In the current logic, deleted links are added to the counter inside CheckNamedRangeLinks
        MsgBox msg & numLinks & " issue(s) detected.", vbExclamation
    End If
End Sub

' ==========================================
' HELPER CHECK PROCEDURES
' ==========================================

Private Sub CheckCellFormulas(wksht As Worksheet, wkbk As Workbook, ByRef numLinks As Long)
    Dim foundCell   As Range
    Dim firstAddress As String
    Dim fml         As String

    On Error Resume Next
    Set foundCell = wksht.UsedRange.Find(What:="[", LookIn:=xlFormulas, LookAt:=xlPart)
    On Error GoTo 0

    If Not foundCell Is Nothing Then
        firstAddress = foundCell.Address
        Do
            On Error Resume Next
            fml = foundCell.Formula
            If Err.Number = 0 Then
                ' Checking for ".xl" helps avoid false positives on text like "[Text]"
                If InStr(1, fml, ".xl", vbTextCompare) > 0 Then
                    numLinks = numLinks + 1
                    Call OutputLinkInfo("Cell Formula", _
                            wkbk.FullName, _
                            wksht.Name, _
                            "Cell " & foundCell.Address(False, False), _
                            foundCell.Address, _
                            fml, _
                            "Edit link on this sheet.")
                End If
            Else
                Err.Clear
            End If
            On Error GoTo 0

            Set foundCell = wksht.UsedRange.FindNext(foundCell)
            If foundCell Is Nothing Then Exit Do
            If foundCell.Address = firstAddress Then Exit Do
        Loop
    End If
End Sub

Private Sub CheckShapeLinks(wksht As Worksheet, wkbk As Workbook, ByRef numLinks As Long)
    Dim shp         As Shape
    Dim subshp      As Shape
    Dim fml         As String

    For Each shp In wksht.Shapes
        On Error Resume Next
        fml = shp.DrawingObject.Formula
        If Err.Number = 0 And InStr(fml, "[") <> 0 Then
            numLinks = numLinks + 1
            Call OutputLinkInfo("Shape/Object", _
                    wkbk.FullName, _
                    wksht.Name, _
                    shp.Name, _
                    shp.TopLeftCell.Address & ":" & shp.BottomRightCell.Address, _
                    fml, _
                    "Delete link. Link can be changed via Excel menu. Delete object.")
        End If
        On Error GoTo 0

        ' Check grouped shapes
        If shp.Type = msoGroup Then
            For Each subshp In shp.GroupItems
                On Error Resume Next
                fml = subshp.DrawingObject.Formula
                If Err.Number = 0 And InStr(fml, "[") <> 0 Then
                    numLinks = numLinks + 1
                    Call OutputLinkInfo("Shape/Object", _
                            wkbk.FullName, _
                            wksht.Name, _
                            subshp.Name & " (part of group '" & shp.Name & "')", _
                            subshp.TopLeftCell.Address & ":" & subshp.BottomRightCell.Address, _
                            fml, _
                            "Delete link. Link can be changed via Excel menu. Delete object.")
                End If
                On Error GoTo 0
            Next subshp
        End If
    Next shp
End Sub

Private Sub CheckConditionalFormatting(wksht As Worksheet, wkbk As Workbook, ByRef numLinks As Long)
    Dim cForm       As FormatCondition
    Dim fml         As String

    For Each cForm In wksht.Cells.FormatConditions
        On Error Resume Next
        fml = cForm.Formula1
        If Err.Number = 0 And InStr(fml, "[") <> 0 Then
            numLinks = numLinks + 1
            Call OutputLinkInfo("Conditional Formatting", _
                    wkbk.FullName, _
                    wksht.Name, _
                    "Cell " & cForm.AppliesTo.Address(False, False), _
                    cForm.AppliesTo.Address, _
                    fml, _
                    "Link found in conditional formatting. Excel often does not display it in the interface." & _
                    "It is recommended to delete conditional formatting for these cells or replace it with a rule without links.")
        End If
        On Error GoTo 0
    Next cForm
End Sub

Private Sub CheckChartLinks(wksht As Worksheet, wkbk As Workbook, ByRef numLinks As Long)
    Dim cht         As ChartObject
    Dim srs         As Series
    Dim chartName   As String
    Dim fml         As String

    For Each cht In wksht.ChartObjects
        For Each srs In cht.Chart.SeriesCollection
            On Error Resume Next
            fml = srs.Formula
            If Err.Number = 0 And InStr(fml, "[") <> 0 Then
                numLinks = numLinks + 1
                If cht.Chart.HasTitle Then
                    chartName = cht.Chart.ChartTitle.Caption
                Else
                    chartName = cht.Chart.Name & " (" & cht.Name & ")"
                End If

                Call OutputLinkInfo("Chart", _
                        wkbk.FullName, _
                        wksht.Name, _
                        chartName, _
                        cht.TopLeftCell.Address & ":" & cht.BottomRightCell.Address, _
                        fml, _
                        "Change chart data source (Select Data). Find the Series with the problematic link.")
            End If
            On Error GoTo 0
        Next srs
    Next cht
End Sub

Private Sub CheckPivotTableLinks(wksht As Worksheet, wkbk As Workbook, ByRef numLinks As Long)
    Dim pvt         As PivotTable
    Dim fml         As String

    For Each pvt In wksht.PivotTables
        On Error Resume Next
        fml = pvt.SourceData
        If Err.Number = 0 And InStr(fml, "[") <> 0 Then
            numLinks = numLinks + 1
            Call OutputLinkInfo("Pivot Table", _
                    wkbk.FullName, _
                    wksht.Name, _
                    pvt.Name, _
                    pvt.TableRange1.Address, _
                    fml, _
                    "External workbook link. Change pivot table data source.")
        End If
        On Error GoTo 0
    Next pvt
End Sub

Private Sub CheckDataValidationLinks(wksht As Worksheet, wkbk As Workbook, ByRef numLinks As Long)
    Dim r           As Range
    Dim cell        As Range
    Dim fml         As String
    Dim dataValExtLinkRanges As Object
    Dim key         As Variant
    Dim contiguousAddresses() As String
    Dim i           As Long
    Dim place       As String

    Set dataValExtLinkRanges = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    Set r = wksht.Cells.SpecialCells(xlCellTypeAllValidation)
    On Error GoTo 0

    If Not r Is Nothing Then
        For Each cell In r.Cells
            On Error Resume Next
            fml = cell.Validation.Formula1
            If Err.Number = 0 And InStr(fml, "[") <> 0 Then
                On Error GoTo 0
                key = fml
                If dataValExtLinkRanges.Exists(key) Then
                    Set dataValExtLinkRanges.item(key) = Application.Union(dataValExtLinkRanges(key), cell)
                Else
                    Set dataValExtLinkRanges.item(key) = cell
                End If
            End If
            On Error GoTo 0
        Next cell
    End If

    ' Output results for data validation
    For Each key In dataValExtLinkRanges.keys
        contiguousAddresses = VBA.Split(dataValExtLinkRanges(key).Address, ",")
        For i = 0 To UBound(contiguousAddresses)
            numLinks = numLinks + 1
            If Range(contiguousAddresses(i)).CountLarge > 1 Then
                place = "Cells " & VBA.Replace(contiguousAddresses(i), "$", "")
            Else
                place = "Cell " & VBA.Replace(contiguousAddresses(i), "$", "")
            End If

            Call OutputLinkInfo("Data Validation", _
                    wkbk.FullName, _
                    wksht.Name, _
                    place, _
                    contiguousAddresses(i), _
                    VBA.CStr(key), _
                    "Link found in data validation (Data -> Data Validation). Change the source.")
        Next i
    Next key

    Set dataValExtLinkRanges = Nothing
End Sub

Private Sub CheckNamedRangeLinks(wkbk As Workbook, ByRef numLinks As Long)
    Dim fso         As Object
    Dim startPos    As Long
    Dim endPos      As Long
    Dim pathPos     As Long
    Dim delCt       As Long
    Dim nameCnt     As Integer
    Dim iCount      As Integer
    Dim FilePath    As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    delCt = 0
    iCount = wkbk.Names.Count

    If iCount > 0 Then
        For nameCnt = iCount To 1 Step -1
            If InStr(wkbk.Names(nameCnt).RefersTo, "[") <> 0 Then
                ' Delete broken links (#REF!)
                If InStr(wkbk.Names(nameCnt).RefersTo, "#REF!") <> 0 Then
                    wkbk.Names(nameCnt).Delete
                    delCt = delCt + 1
                Else
                    ' Check if file exists
                    startPos = VBA.InStr(1, wkbk.Names(nameCnt).RefersTo, "='")
                    If startPos > 0 Then
                        endPos = VBA.InStr(startPos, wkbk.Names(nameCnt).RefersTo, "]")
                        pathPos = VBA.InStr(1, wkbk.Names(nameCnt).RefersTo, "\")

                        If startPos > 0 And endPos > 0 And pathPos > 0 Then
                            FilePath = VBA.Replace(VBA.mid(wkbk.Names(nameCnt).RefersTo, startPos + 2, endPos - startPos - 2), "[", "")
                            If Not fso.FileExists(FilePath) Then
                                wkbk.Names(nameCnt).Delete
                                delCt = delCt + 1
                            Else
                                ' File exists, but link is external
                                wkbk.Names(nameCnt).Visible = True
                                numLinks = numLinks + 1
                                Call OutputLinkInfo("Named Range", _
                                        wkbk.FullName, _
                                        "N/A", _
                                        wkbk.Names(nameCnt).Name, _
                                        "", _
                                        wkbk.Names(nameCnt).RefersTo, _
                                        "Open Name Manager (Formulas -> Name Manager). Check the path.")
                            End If
                        End If
                    End If
                End If
            End If
        Next nameCnt
    End If

    Set fso = Nothing

    ' Report on deleted names
    If delCt > 0 Then
        numLinks = numLinks + 1
        Call OutputLinkInfo("Named Range", _
                wkbk.FullName, _
                "N/A", _
                "(" & delCt & " deleted names)", _
                "", _
                "Not recorded", _
                "Number of deleted named ranges with broken links. File: " & Dir(wkbk.FullName))
    End If
End Sub