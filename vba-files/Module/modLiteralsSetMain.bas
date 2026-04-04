Attribute VB_Name = "modLiteralsSetMain"
Option Explicit
Option Private Module

Public Sub ReNameLiteralsFile()

    Dim wbSetings   As Workbook
    If Not GetTargetWorkbook(wbSetings, "Renaming String Literals:", "RENAME") Then Exit Sub

    Application.ScreenUpdating = False
    Dim arrCode     As Variant
    Dim arrUserForm As Variant
    Dim arrUI       As Variant
    Dim sPath       As String

    Const MSG_FILE_NOT_FOUND As String = "File not found: "
    Const MSG_PROJECT_LOCKED As String = "VBA project is password protected!"

    Call getArrayFromSheet(wbSetings, sPath, STR_CODE, arrCode, 4)
    Call getArrayFromSheet(wbSetings, sPath, STR_UF, arrUserForm, 8)
    Call getArrayFromSheet(wbSetings, sPath, STR_UI, arrUI, 9)

    If Not FileHave(sPath, vbNormal) Then
        Call MsgBox(MSG_FILE_NOT_FOUND & sPath, vbCritical, "Error")
        Exit Sub
    End If

    Dim wb          As Workbook
    Dim vbProj      As VBIDE.vbProject
    Dim sNameFile   As String
    sNameFile = sGetFileName(sPath)
    Set wb = getWorkBook(sNameFile)
    If wb Is Nothing Then
        Application.DisplayAlerts = False
        Set wb = Workbooks.Open(FileName:=sPath, UpdateLinks:=0)
        Application.DisplayAlerts = True
    End If
    Set vbProj = wb.vbProject

    If vbProj.Protection = vbext_pp_locked Then
        Call MsgBox(MSG_PROJECT_LOCKED, vbCritical, "Access denied")
        Exit Sub
    End If

    Dim dtStart     As Date
    dtStart = VBA.Now()
    Debug.Print ">> Data collection: " & wb.FullName

    If renameLiteralsToUserForm(vbProj, arrUserForm) Then
        Call loadArrayToSheet(wbSetings, STR_UF, arrUserForm)
        Debug.Print vbTab & ">> " & VBA.Format$(VBA.Now() - dtStart, FORMAT_TIME) & " User Forms"
    End If

    If renameLiteralsToCode(vbProj, arrCode) Then
        Call loadArrayToSheet(wbSetings, STR_CODE, arrCode)
        Debug.Print vbTab & ">> " & VBA.Format$(VBA.Now() - dtStart, FORMAT_TIME) & " Code VBA"
    End If

    If renameLiteralsToUI(wb, arrUI) Then
        Call loadArrayToSheet(wbSetings, STR_UI, arrUI)
        Debug.Print vbTab & ">> " & VBA.Format$(VBA.Now() - dtStart, FORMAT_TIME) & " Ribbon Control UI"
    End If
End Sub

Private Sub loadArrayToSheet(ByRef wb As Workbook, ByRef sNameSheet As String, ByRef arr As Variant)
    If IsEmpty(arr) Then Exit Sub
    With wb.Worksheets(sNameSheet)
        .Cells(3, 1).Resize(UBound(arr, 1), UBound(arr, 2)).Value2 = arr
    End With
End Sub

Private Sub getArrayFromSheet(ByRef wb As Workbook, ByRef sPath As String, ByRef sNameSheet As String, ByRef arr As Variant, ByRef iColEnd As Integer)
    Dim sh          As Worksheet
    Set sh = getWorkSheet(wb, sNameSheet)
    If sh Is Nothing Then Exit Sub
    With sh
        Dim lLastRow As Long
        lLastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        If lLastRow < 3 Then Exit Sub
        If VBA.Len(sPath) = 0 Then sPath = .Cells(1, 1).Value2
        arr = .Range(.Cells(3, 1), .Cells(lLastRow, iColEnd)).Value2
    End With
End Sub

Private Function getWorkSheet(ByRef wb As Workbook, ByRef sNameFile As String) As Worksheet
    On Error Resume Next
    Set getWorkSheet = wb.Worksheets(sNameFile)
    On Error GoTo 0
End Function

Private Function getWorkBook(ByRef sNameFile As String) As Workbook
    On Error Resume Next
    Set getWorkBook = Workbooks(sNameFile)
    On Error GoTo 0
End Function